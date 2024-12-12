/*
 * Copyright 2016 - 2021 Draco, https://github.com/draco1023
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ddr.poi.html.tag;

import com.drew.imaging.FileType;
import com.drew.imaging.FileTypeDetector;
import com.drew.imaging.ImageMetadataReader;
import com.drew.imaging.ImageProcessingException;
import com.drew.metadata.Metadata;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.SVGPictureData;
import org.ddr.image.ImageInfo;
import org.ddr.image.ImageType;
import org.ddr.image.MetadataReader;
import org.ddr.image.MetadataReaders;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.html.util.CSSLength;
import org.ddr.poi.math.MathMLUtils;
import org.ddr.poi.util.ByteArrayCopyStream;
import org.ddr.poi.util.HttpURLConnectionUtils;
import org.jsoup.nodes.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.Iterator;

/**
 * img标签渲染器
 *
 * @author Draco
 * @since 2021-02-09
 */
public class ImageRenderer implements ElementRenderer {
    private static final Logger log = LoggerFactory.getLogger(ImageRenderer.class);

    private static final String[] TAGS = {HtmlConstants.TAG_IMG};
    private static final String HTTP = "http";
    private static final String DOUBLE_SLASH = "//";
    private static final String DATA_PREFIX = "data:";
    private static final String COMMENT_MATH_PREFIX = "<!--MathML: <math ";
    private static final String COMMENT_MATH_SUFFIX = "</math>-->";

    static {
        SVGPictureData.initRelation();
    }

    /**
     * 元素渲染结束需要执行的逻辑
     *
     * @param element HTML元素
     * @param context 渲染上下文
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        String src = element.attr(HtmlConstants.ATTR_SRC);
        if (StringUtils.startsWithIgnoreCase(src, HTTP)) {
            handleRemoteImage(element, context, src);
        } else if (StringUtils.startsWith(src, DOUBLE_SLASH)) {
            // 某些图片链接为了跟随网站协议而隐去了协议名称
            handleRemoteImage(element, context, HTTP + HtmlConstants.COLON + src);
        } else if (StringUtils.startsWith(src, DATA_PREFIX)) {
            handleData(element, context, src);
        }
        return false;
    }

    /**
     * 处理Data URL
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @param src 数据
     */
    protected void handleData(Element element, HtmlRenderContext context, String src) {
        int index = src.indexOf(HtmlConstants.COMMA.charAt(0));
        String data = src.substring(index + 1);
        String declaration = src.substring(0, index);
        String format = StringUtils.substringBetween(declaration, HtmlConstants.SLASH, HtmlConstants.SEMICOLON);
        // org.apache.poi.sl.usermodel.PictureData.PictureType
        if (format.contains(HtmlConstants.MINUS)) {
            format = StringUtils.substringAfterLast(format, HtmlConstants.MINUS);
        } else if (format.contains(HtmlConstants.PLUS)) {
            format = StringUtils.substringBefore(format, HtmlConstants.PLUS);
        }

        byte[] bytes;
        if (declaration.contains("base64")) {
            try {
                bytes = Base64.getDecoder().decode(data);
            } catch (Exception e) {
                log.warn("Failed to load image due to illegal base64 data: {}", src);
                return;
            }
        } else {
            if (data.startsWith(HtmlConstants.PERCENT)) {
                try {
                    data = URLDecoder.decode(data, StandardCharsets.UTF_8.name());
                } catch (UnsupportedEncodingException e) {
                    log.warn("Failed to load image due to illegal data url: {}", src);
                    return;
                }
            }

            // wiris support
            int startOfMath = data.indexOf(COMMENT_MATH_PREFIX);
            if (startOfMath >= 0) {
                try {
                    int endOfMath = data.indexOf(COMMENT_MATH_SUFFIX, startOfMath + COMMENT_MATH_PREFIX.length());
                    String math = data.substring(startOfMath + 12, endOfMath + 7);
                    MathMLUtils.renderTo(context.getClosestParagraph(), context.newRun(), math);
                    return;
                } catch (Exception e) {
                    log.warn("Failed to render math in wiris svg, will try to render as svg image: {}", data, e);
                }
            }

            bytes = data.getBytes(StandardCharsets.UTF_8);
        }
        boolean svg = HtmlConstants.TAG_SVG.equals(format);
        try (ByteArrayInputStream inputStream = new ByteArrayInputStream(bytes)) {
            ImageInfo info = analyzeImage(inputStream, svg);
            if (info == null) {
                log.warn("Illegal image url: {}", src);
                return;
            }
            addPicture(element, context, info.getStream(), info.getRawType(), info.getWidth(), info.getHeight(), svg ? bytes : null);
        } catch (IOException | InvalidFormatException e) {
            log.warn("Failed to load image: {}", src, e);
        }
    }

    /**
     * 根据图片反推类型
     *
     * @param image 图片
     * @return 图片类型
     */
    protected ImageType typeOf(BufferedImage image) {
        return image.getColorModel().hasAlpha() ? ImageType.PNG : ImageType.JPG;
    }

    /**
     * 处理远程图片
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @param src 图片链接地址
     */
    protected void handleRemoteImage(Element element, HtmlRenderContext context, String src) {
        HttpURLConnection connect = null;
        try {
            connect = HttpURLConnectionUtils.connect(src);
            InputStream urlStream = connect.getInputStream();
            boolean svg = StringUtils.contains(connect.getHeaderField("content-type"), HtmlConstants.TAG_SVG);
            ByteArrayCopyStream outputStream = new ByteArrayCopyStream(urlStream.available());
            IOUtils.copy(urlStream, outputStream);
            final byte[] svgData = svg ? outputStream.toByteArray() : null;

            ByteArrayInputStream inputStream = outputStream.toInput();
            ImageInfo info = analyzeImage(inputStream, svg);
            if (info == null) {
                log.warn("Illegal image url: {}", src);
                return;
            }

            addPicture(element, context, info.getStream(), info.getRawType(), info.getWidth(), info.getHeight(), svgData);
        } catch (IOException | InvalidFormatException e) {
            log.warn("Failed to load image: {}", src, e);
        } finally {
            IOUtils.close(connect);
        }
    }

    protected ImageInfo analyzeImage(ByteArrayInputStream inputStream, boolean svg) throws IOException, InvalidFormatException {
        final long length = inputStream.available();
        // actual image data stream
        ByteArrayInputStream stream = inputStream;
        ImageType type = null;
        Dimension dimension = null;

        if (svg) {
            BufferedImage image = ImageIO.read(inputStream);
            inputStream.reset();

            type = typeOf(image);
            ByteArrayCopyStream imageStream = new ByteArrayCopyStream(image.getData().getDataBuffer().getSize());
            ImageIO.write(image, type.getExtension(), imageStream);
            stream = imageStream.toInput();

            dimension = new Dimension(image.getWidth(), image.getHeight());
        } else {
            FileType fileType = FileTypeDetector.detectFileType(inputStream);
            for (MetadataReader metadataReader : MetadataReaders.INSTANCES) {
                if (metadataReader.canRead(fileType)) {
                    try {
                        Metadata metadata = ImageMetadataReader.readMetadata(inputStream, length, fileType);
                        type = metadataReader.getType(metadata);
                        dimension = metadataReader.getDimension(metadata);
                        break;
                    } catch (ImageProcessingException ignored) {
                    }
                }
            }
            inputStream.reset();
            if (dimension == null) {
                Iterator<ImageReader> imageReaders = ImageIO.getImageReaders(inputStream);
                while (imageReaders.hasNext()) {
                    ImageReader reader = imageReaders.next();
                    try {
                        dimension = new Dimension(reader.getWidth(0), reader.getHeight(0));
                        break;
                    } catch (IOException ignored) {
                    }
                }
                if (dimension == null) {
                    BufferedImage image = ImageIO.read(inputStream);
                    inputStream.reset();

                    if (image == null) {
                        return null;
                    }

                    if (type == null) {
                        type = typeOf(image);
                    }
                    dimension = new Dimension(image.getWidth(), image.getHeight());
                }
            }
        }
        return new ImageInfo(stream, type, dimension);
    }

    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    @Override
    public boolean renderAsBlock() {
        return false;
    }

    /**
     * 添加图片
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @param inputStream 图片数据流
     * @param type 图片类型
     * @param widthInPixels 图片宽度（像素）
     * @param heightInPixels 图片高度（像素）
     * @param svgData SVG数据
     */
    protected void addPicture(Element element, HtmlRenderContext context, InputStream inputStream, int type,
                              int widthInPixels, int heightInPixels,
                              byte[] svgData) throws InvalidFormatException, IOException {
        // 容器限制宽度
        int containerWidth = context.getAvailableWidthInEMU();
//        int containerHeight = context.getAvailablePageHeight();
        // 图片原始宽高
        int widthInEMU = Units.pixelToEMU(widthInPixels);
        int heightInEMU = Units.pixelToEMU(heightInPixels);
        float naturalAspect = 1f * widthInEMU / heightInEMU;

        int declaredWidth = widthInEMU;
        int declaredHeight = heightInEMU;
        int maxWidthInEMU = containerWidth;
        int maxHeightInEMU = Integer.MAX_VALUE;

        String width = context.getPropertyValue(HtmlConstants.CSS_WIDTH);
        if (width.length() > 0) {
            CSSLength cssLength = CSSLength.of(width);
            if (cssLength.isValid()) {
                declaredWidth = context.computeLengthInEMU(cssLength, widthInEMU, containerWidth);
            }
        } else {
            // width attribute is overridden by style, the same to height
            // https://css-tricks.com/whats-the-difference-between-width-height-in-css-and-width-height-html-attributes/
            width = element.attr(HtmlConstants.ATTR_WIDTH);
            if (NumberUtils.isParsable(width)) {
                width += HtmlConstants.PX;
                CSSLength cssLength = CSSLength.of(width);
                declaredWidth = context.computeLengthInEMU(cssLength, widthInEMU, containerWidth);
            }
        }


        String maxWidth = context.getPropertyValue(HtmlConstants.CSS_MAX_WIDTH);
        if (maxWidth.length() > 0) {
            CSSLength cssLength = CSSLength.of(maxWidth);
            if (cssLength.isValid()) {
                // restrained by container
                maxWidthInEMU = Math.min(context.computeLengthInEMU(cssLength, widthInEMU, containerWidth), containerWidth);
            }
        }

        String height = context.getPropertyValue(HtmlConstants.CSS_HEIGHT);
        if (height.length() > 0) {
            CSSLength cssLength = CSSLength.of(height);
            if (cssLength.isValid()) {
                declaredHeight = context.computeLengthInEMU(cssLength, heightInEMU, Integer.MAX_VALUE);
            }
        } else {
            height = element.attr(HtmlConstants.ATTR_HEIGHT);
            if (NumberUtils.isParsable(height)) {
                height += HtmlConstants.PX;
                CSSLength cssLength = CSSLength.of(height);
                declaredHeight = context.computeLengthInEMU(cssLength, heightInEMU, Integer.MAX_VALUE);
            }
        }

        String maxHeight = context.getPropertyValue(HtmlConstants.CSS_MAX_HEIGHT);
        if (maxHeight.length() > 0) {
            CSSLength cssLength = CSSLength.of(maxHeight);
            if (cssLength.isValid()) {
                maxHeightInEMU = context.computeLengthInEMU(cssLength, heightInEMU, Integer.MAX_VALUE);
            }
        }

        if (declaredWidth == widthInEMU ^ declaredHeight == heightInEMU) {
            if (declaredWidth == widthInEMU) {
                declaredWidth = (int) (declaredHeight * naturalAspect);
            } else {
                declaredHeight = (int) (declaredWidth / naturalAspect);
            }
        }

        // 计算尺寸
        int calculatedWidth, calculatedHeight;
        if (declaredWidth < maxWidthInEMU && declaredHeight <= maxHeightInEMU) {
            calculatedWidth = declaredWidth;
            calculatedHeight = declaredHeight;
        } else if (declaredWidth > maxWidthInEMU && declaredHeight <= maxHeightInEMU) {
            calculatedWidth = maxWidthInEMU;
            calculatedHeight = (int) (maxWidthInEMU / naturalAspect);
        } else if (declaredHeight > maxHeightInEMU && declaredWidth <= maxWidthInEMU) {
            calculatedHeight = maxHeightInEMU;
            calculatedWidth = (int) (maxHeightInEMU * naturalAspect);
        } else {
            float widthRatio = 1f * maxWidthInEMU / declaredWidth;
            float heightRatio = 1f * maxHeightInEMU / declaredHeight;
            float scale = Math.min(widthRatio, heightRatio);
            calculatedWidth = (int) (declaredWidth * scale);
            calculatedHeight = (int) (declaredHeight * scale);
        }

        context.renderPicture(inputStream, type, HtmlConstants.TAG_IMG,
            calculatedWidth, calculatedHeight, svgData);
    }

}
