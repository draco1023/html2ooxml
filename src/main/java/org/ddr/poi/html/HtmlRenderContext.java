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

package org.ddr.poi.html;

import com.deepoove.poi.render.RenderContext;
import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ooxml.util.POIXMLUnits;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BodyType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.SVGPictureData;
import org.apache.poi.xwpf.usermodel.SVGRelation;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFFootnote;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.ddr.poi.html.util.CSSLength;
import org.ddr.poi.html.util.CSSLengthUnit;
import org.ddr.poi.html.util.CSSStyleUtils;
import org.ddr.poi.html.util.Colors;
import org.ddr.poi.html.util.InlineStyle;
import org.ddr.poi.html.util.NamedFontSize;
import org.ddr.poi.html.util.NumberingContext;
import org.ddr.poi.html.util.RenderUtils;
import org.ddr.poi.html.util.WhiteSpaceRule;
import org.ddr.poi.html.util.XWPFParagraphRuns;
import org.ddr.poi.util.XmlUtils;
import org.jsoup.internal.StringUtil;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectFrameLocking;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualGraphicFrameProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTOfficeArtExtension;
import org.openxmlformats.schemas.drawingml.x2006.main.CTOfficeArtExtensionList;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTPosH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTPosV;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignV;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromV;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STVerticalAlignRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STThemeColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URI;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

/**
 * HTML字符串渲染上下文
 *
 * @author Draco
 * @since 2021-02-08
 */
public class HtmlRenderContext extends RenderContext<String> {
    private static final Logger log = LoggerFactory.getLogger(HtmlRenderContext.class);


    /**
     * 默认字号 小四 12pt 16px
     */
    private static final CSSLength DEFAULT_FONT_SIZE = new CSSLength(12, CSSLengthUnit.PT);
    /**
     * 默认超链接颜色
     */
    private static final String DEFAULT_HYPERLINK_COLOR = "0563C1";

    /**
     * HTML元素渲染器提供者
     */
    private final ElementRendererProvider rendererProvider;

    /**
     * 父容器（子元素通常为段落/表格）栈，主要用于渲染HTML表格时父容器的切换
     */
    private LinkedList<IBody> ancestors = new LinkedList<>();

    /**
     * 行内样式栈，最近声明的样式最先生效
     */
    private LinkedList<InlineStyle> inlineStyles = new LinkedList<>();

    /**
     * 字号栈，一些相对大小的字号值将被进行换算
     */
    private LinkedList<Integer> fontSizesInHalfPoints = new LinkedList<>();

    /**
     * 列表上下文，用于处理嵌套列表
     */
    private final NumberingContext numberingContext;
    /**
     * 默认字号
     */
    private final CSSLength defaultFontSize;
    /**
     * 页面宽度
     */
    private final CSSLength pageWidth;
    /**
     * 页面高度
     */
    private final CSSLength pageHeight;
    /**
     * 页面顶部边距
     */
    private final CSSLength marginTop;
    /**
     * 页面右侧边距
     */
    private final CSSLength marginRight;
    /**
     * 页面底部边距
     */
    private final CSSLength marginBottom;
    /**
     * 页面左侧边距
     */
    private final CSSLength marginLeft;
    /**
     * 可用页面宽度
     */
    private final int availablePageWidth;
    /**
     * 可用页面高度
     */
    private final int availablePageHeight;
    /**
     * 占位符所在段落的样式ID
     */
    private String placeholderStyleId;

    /**
     * 当前Run元素，可能为超链接
     */
    private XWPFRun currentRun;

    /**
     * 全局字体，声明后所有样式中的字体将失效
     */
    private String globalFont;
    /**
     * 全局字号，声明后所有样式中的字号将失效
     */
    private BigInteger globalFontSize;
    /**
     * 嵌套表格是否默认显示边框
     */
    private boolean showDefaultTableBorderInTableCell;

    /**
     * 块状元素深度计数器
     */
    private int blockLevel;

    /**
     * 全局xml指针，保证其总是处于将要插入内容的位置，仅在需要移动之前进行push，适时pop还原位置
     */
    private XmlCursor globalCursor;

    /**
     * 防重段落
     */
    private XWPFParagraph dedupeParagraph;

    /**
     * 同一段落内的前一个文本节点
     */
    private TextWrapper previousText;

    /**
     * 构造方法
     *
     * @param context 原始渲染上下文
     */
    public HtmlRenderContext(RenderContext<String> context, ElementRendererProvider rendererProvider) {
        super(context.getEleTemplate(), context.getData(), context.getTemplate());
        this.rendererProvider = rendererProvider;
        globalCursor = getRun().getCTR().newCursor();

        numberingContext = new NumberingContext(getXWPFDocument());

        long w = RenderUtils.A4_WIDTH * Units.EMU_PER_DXA;
        long h = RenderUtils.A4_HEIGHT * Units.EMU_PER_DXA;
        long top = RenderUtils.DEFAULT_TOP_MARGIN * Units.EMU_PER_DXA;
        long right = RenderUtils.DEFAULT_RIGHT_MARGIN * Units.EMU_PER_DXA;
        long bottom = RenderUtils.DEFAULT_BOTTOM_MARGIN * Units.EMU_PER_DXA;
        long left = RenderUtils.DEFAULT_LEFT_MARGIN * Units.EMU_PER_DXA;
        CTSectPr sectPr = getXWPFDocument().getDocument().getBody().getSectPr();
        if (sectPr != null) {
            CTPageSz pgSz = sectPr.getPgSz();

            if (pgSz != null) {
                w = POIXMLUnits.parseLength(pgSz.xgetW());
                h = POIXMLUnits.parseLength(pgSz.xgetH());
            }

            CTPageMar pgMar = sectPr.getPgMar();
            if (pgMar != null) {
                top = POIXMLUnits.parseLength(pgMar.xgetTop());
                right = POIXMLUnits.parseLength(pgMar.xgetRight());
                bottom = POIXMLUnits.parseLength(pgMar.xgetBottom());
                left = POIXMLUnits.parseLength(pgMar.xgetLeft());
            }
        }

        pageWidth = new CSSLength(w, CSSLengthUnit.EMU);
        pageHeight = new CSSLength(h, CSSLengthUnit.EMU);
        marginTop = new CSSLength(top, CSSLengthUnit.EMU);
        marginRight = new CSSLength(right, CSSLengthUnit.EMU);
        marginBottom = new CSSLength(bottom, CSSLengthUnit.EMU);
        marginLeft = new CSSLength(left, CSSLengthUnit.EMU);
        availablePageWidth = (int) (w - left - right);
        availablePageHeight = (int) (h - top - bottom);

        Double fontSize = getXWPFDocument().getStyles().getDefaultRunStyle().getFontSizeAsDouble();
        defaultFontSize = fontSize != null ? new CSSLength(fontSize, CSSLengthUnit.PT) : DEFAULT_FONT_SIZE;

        extractPlaceholderStyle();
    }

    /**
     * 抽取占位符所在段落的样式
     */
    private void extractPlaceholderStyle() {
        XWPFRun run = getRun();
        IRunBody runParent = run.getParent();
        if (runParent instanceof XWPFParagraph) {
            XWPFParagraph paragraph = (XWPFParagraph) runParent;
            String styleId = paragraph.getStyleID();
            boolean existsRPr = run.getCTR().isSetRPr();

            if (styleId == null && !existsRPr) {
                return;
            } else if (styleId != null && !existsRPr) {
                placeholderStyleId = styleId;
                return;
            }

            XWPFStyles styles = getXWPFDocument().getStyles();
            CTStyle newCTStyle = CTStyle.Factory.newInstance();
            newCTStyle.setCustomStyle(true);
            newCTStyle.setType(STStyleType.PARAGRAPH);
            newCTStyle.addNewHidden();
            newCTStyle.setRPr(run.getCTR().getRPr());
            XmlUtils.removeNamespaces(newCTStyle.getRPr());

            String newStyleId = styleId + styles.getNumberOfStyles();
            newCTStyle.setStyleId(newStyleId);
            newCTStyle.addNewName().setVal(newStyleId);
            placeholderStyleId = newStyleId;

            if (styleId != null) {
                newCTStyle.addNewBasedOn().setVal(styleId);
            }

            XWPFStyle newStyle = new XWPFStyle(newCTStyle, styles);
            styles.addStyle(newStyle);

            paragraph.setStyle(newStyleId);
        }
    }

    @Override
    public IBody getContainer() {
        IBody container = ancestors.peek();
        return container == null ? super.getContainer() : container;
    }

    /**
     * 父容器入栈
     *
     * @param body 父容器
     */
    public void pushContainer(IBody body) {
        ancestors.push(body);
    }

    /**
     * 父容器出栈
     */
    public void popContainer() {
        ancestors.pop();
    }

    /**
     * 获取最近的段落，如果当前最近位置的内容元素是表格，则创建一个与之平级的段落
     *
     * @return 最近的段落
     */
    public XWPFParagraph getClosestParagraph() {
        if (globalCursor.getObject() == getRun().getCTR()) {
            return (XWPFParagraph) getRun().getParent();
        }

        globalCursor.push();
        XWPFParagraph paragraph = null;
        if (globalCursor.toPrevSibling()) {
            XmlObject object = globalCursor.getObject();
            if (object instanceof CTP) {
                paragraph = getContainer().getParagraph((CTP) object);
            } else {
                // pop() is safer than toNextSibling()
                globalCursor.pop();
                globalCursor.push();
                paragraph = newParagraph(null, globalCursor);
                RenderUtils.paragraphStyle(this, paragraph, CSSStyleUtils.EMPTY_STYLE);
            }
        }
        globalCursor.pop();

        if (paragraph != null) {
            return paragraph;
        }

        throw new IllegalStateException("No paragraph in stack");
    }

    /**
     * 开始渲染超链接
     *
     * @param uri 链接地址
     */
    public void startHyperlink(String uri) {
        try {
            URI.create(uri);
        } catch (Exception e) {
            log.warn("Illegal href", e);
            uri = "#";
        }
        if (isBlocked()) {
            XWPFParagraph paragraph = getClosestParagraph();
            currentRun = paragraph.createHyperlinkRun(uri);
            if (dedupeParagraph == paragraph) {
                unmarkDedupe();
            }
        } else {
            // 在占位符之前插入超链接
            String rId = getRun().getParent().getPart().getPackagePart()
                    .addExternalRelationship(uri, XWPFRelation.HYPERLINK.getRelation()).getId();
            XmlCursor xmlCursor = getRun().getCTR().newCursor();
            xmlCursor.insertElement(XmlUtils.HYPERLINK_QNAME);
            xmlCursor.toPrevSibling();
            CTHyperlink ctHyperlink = (CTHyperlink) xmlCursor.getObject();
            xmlCursor.dispose();
            ctHyperlink.setId(rId);
            ctHyperlink.addNewR();
            currentRun = new XWPFHyperlinkRun(ctHyperlink, ctHyperlink.getRArray(0), getRun().getParent());
        }
    }

    /**
     * 结束渲染超链接
     */
    public void endHyperlink() {
        currentRun = null;
    }

    /**
     * 新建段落
     *
     * @param container 容器
     * @param cursor xml指针
     * @return 段落
     */
    public XWPFParagraph newParagraph(IBody container, XmlCursor cursor) {
        if (container == null) {
            container = getContainer();
        }
        XWPFParagraph xwpfParagraph = container.insertNewParagraph(cursor);
        if (placeholderStyleId != null) {
            xwpfParagraph.setStyle(placeholderStyleId);
        }
        markDedupe(xwpfParagraph);
        previousText = null;
        return xwpfParagraph;
    }

    /**
     * 新建CTR
     *
     * @return CTR
     */
    public CTR newRun() {
        // 超链接虽然不是段落，但是内部可以容纳多个run
        if (currentRun instanceof XWPFHyperlinkRun) {
            XmlCursor xmlCursor = currentRun.getCTR().newCursor();
            CTR ctr;
            if (xmlCursor.toFirstChild()) {
                ctr = ((XWPFHyperlinkRun) currentRun).getCTHyperlink().addNewR();
            } else {
                // run没有内容则直接复用
                ctr = currentRun.getCTR();
            }
            xmlCursor.dispose();
            // 默认链接样式
            initHyperlinkStyle(ctr);

            return ctr;
        }
        // 考虑到样式可能不一致，总是创建新的run
        if (isBlocked()) {
            XWPFParagraph paragraph = getClosestParagraph();
            currentRun = paragraph.createRun();
            if (dedupeParagraph == paragraph) {
                unmarkDedupe();
            }
        } else {
            // 在占位符之前插入run
            XmlCursor xmlCursor = getRun().getCTR().newCursor();
            xmlCursor.insertElement(XmlUtils.R_QNAME);
            xmlCursor.toPrevSibling();
            CTR ctr = (CTR) xmlCursor.getObject();
            xmlCursor.dispose();
            currentRun = new XWPFRun(ctr, getRun().getParent());
        }
        return currentRun.getCTR();
    }

    /**
     * 初始化超链接样式
     *
     * @param ctr CTR
     */
    private void initHyperlinkStyle(CTR ctr) {
        CTRPr rPr = RenderUtils.getRPr(ctr);
        CTColor ctColor = rPr.addNewColor();
        ctColor.setVal(DEFAULT_HYPERLINK_COLOR);
        ctColor.setThemeColor(STThemeColor.HYPERLINK);

        rPr.addNewU().setVal(STUnderline.SINGLE);
    }

    /**
     * 获取最近的表格，仅可在渲染表格及其内部元素的时候使用
     *
     * @return 最近的表格
     */
    public XWPFTable getClosestTable() {
        globalCursor.push();
        XWPFTable table = null;
        if (globalCursor.toPrevSibling()) {
            XmlObject object = globalCursor.getObject();
            if (object instanceof CTTbl) {
                table = getContainer().getTable((CTTbl) object);
            }
        }
        globalCursor.pop();

        if (table != null) {
            return table;
        }

        throw new IllegalStateException("No table in stack");
    }

    /**
     * 行内样式入栈
     *
     * @param inlineStyle 样式声明
     * @param block 是否为块状元素
     */
    public void pushInlineStyle(CSSStyleDeclarationImpl inlineStyle, boolean block) {
        String newFontSize = inlineStyle.getFontSize();
        // 默认值表示未声明字号
        int fontSize = Integer.MIN_VALUE;
        if (StringUtils.isNotBlank(newFontSize)) {
            NamedFontSize namedFontSize = NamedFontSize.of(newFontSize);
            if (namedFontSize != null) {
                // 固定名称的字号
                fontSize = namedFontSize.getSize().toHalfPoints();
            } else if (HtmlConstants.SMALLER.equalsIgnoreCase(newFontSize)) {
                // 相对小一号
                int inheritedFontSize = getInheritedFontSizeInHalfPoints();
                fontSize = RenderUtils.smallerFontSizeInHalfPoints(inheritedFontSize);
            } else if (HtmlConstants.LARGER.equalsIgnoreCase(newFontSize)) {
                // 相对大一号
                int inheritedFontSize = getInheritedFontSizeInHalfPoints();
                fontSize = RenderUtils.largerFontSizeInHalfPoints(inheritedFontSize);
            } else {
                CSSLength cssLength = CSSLength.of(newFontSize);
                if (cssLength.isValid()) {
                    if (cssLength.getUnit() == CSSLengthUnit.PERCENT) {
                        fontSize = (int) Math.rint(getInheritedFontSizeInHalfPoints()
                                * cssLength.getValue() * cssLength.getUnit().absoluteFactor());
                    } else {
                        int emu = lengthToEMU(cssLength);
                        fontSize = emu * 2 / Units.EMU_PER_POINT;
                    }
                }
            }
        }
        fontSizesInHalfPoints.push(fontSize);

        // text-decoration-line 在继承时需要合并
        String textDecorationLine = inlineStyle.getPropertyValue(HtmlConstants.CSS_TEXT_DECORATION_LINE);
        if (StringUtils.isNotBlank(textDecorationLine) && !HtmlConstants.NONE.equals(textDecorationLine)) {
            Set<String> remainValues = new HashSet<>(HtmlConstants.TEXT_DECORATION_LINES);
            String[] values = StringUtils.split(textDecorationLine, ' ');
            for (String value : values) {
                remainValues.remove(value);
            }

            if (!remainValues.isEmpty()) {
                StringBuilder lines = new StringBuilder(textDecorationLine);
                for (InlineStyle inheritedStyle : inlineStyles) {
                    String s = inheritedStyle.getDeclaration().getPropertyValue(HtmlConstants.CSS_TEXT_DECORATION_LINE);
                    if (HtmlConstants.NONE.equals(s)) {
                        break;
                    } else if (remainValues.contains(s)) {
                        lines.append(' ').append(s);
                        remainValues.remove(s);
                        if (remainValues.isEmpty()) {
                            break;
                        }
                    }
                }
                if (lines.length() > textDecorationLine.length()) {
                    inlineStyle.setProperty(HtmlConstants.CSS_TEXT_DECORATION_LINE, lines.toString(), null);
                }
            }
        }

        inlineStyles.push(new InlineStyle(inlineStyle, block));
    }

    /**
     * 行内样式出栈
     */
    public void popInlineStyle() {
        fontSizesInHalfPoints.pop();
        inlineStyles.pop();
    }

    /**
     * 当前元素的样式声明，在HTML元素渲染开始时立即调用才可得到正确的声明，因为在解析的过程中可能会动态插入样式
     *
     * @return 当前元素的样式声明
     */
    public CSSStyleDeclarationImpl currentElementStyle() {
        InlineStyle inlineStyle = inlineStyles.peek();
        return inlineStyle == null ? CSSStyleUtils.EMPTY_STYLE : inlineStyle.getDeclaration();
    }

    /**
     * 获取样式值，将被转换为小写
     *
     * @param property 样式名称
     * @return 样式值，未声明时返回空字符串
     */
    public String getPropertyValue(String property) {
        return getPropertyValue(property, false);
    }

    /**
     * 获取样式值，将被转换为小写
     *
     * @param property 样式名称
     * @param inlineOnly 是否仅获取行内元素样式
     * @return 样式值，未声明时返回空字符串
     */
    public String getPropertyValue(String property, boolean inlineOnly) {
        return getPropertyValue(property, false, inlineOnly);
    }

    /**
     * 获取样式值
     *
     * @param property 样式名称
     * @param caseSensitive 是否大小写无关，如果无关则将转换为小写，否则保留原始值
     * @param inlineOnly 是否仅获取行内元素样式
     * @return 样式值，未声明时返回空字符串
     */
    public String getPropertyValue(String property, boolean caseSensitive, boolean inlineOnly) {
        for (InlineStyle inlineStyle : inlineStyles) {
            if (inlineOnly && inlineStyle.isBlock()) {
                break;
            }
            String propertyValue = inlineStyle.getDeclaration().getPropertyValue(property);
            if (StringUtils.isNotBlank(propertyValue)) {
                return caseSensitive ? propertyValue : propertyValue.toLowerCase();
            }
        }
        return "";
    }

    /**
     * @return Word中设置的默认字号
     */
    public CSSLength getDefaultFontSize() {
        return defaultFontSize;
    }

    /**
     * @return 获取当前元素继承的字号，以“半点”为单位
     */
    public int getInheritedFontSizeInHalfPoints() {
        for (Integer fontSize : fontSizesInHalfPoints) {
            if (fontSize > 0) {
                return fontSize;
            }
        }
        return defaultFontSize.toHalfPoints();
    }

    /**
     * @return 父容器的可用宽度，以EMU为单位
     */
    public int getAvailableWidthInEMU() {
        IBody container = getContainer();
        if (container.getPartType() == BodyType.DOCUMENT) {
            return availablePageWidth;
        } else {
            return RenderUtils.getAvailableWidthInEMU(container);
        }
    }

    /**
     * 考虑约束计算长度，以EMU为单位
     *
     * @param length 长度声明
     * @param maxLength 最大长度声明
     * @param naturalEMU 原始长度
     * @param parentEMU 父容器长度
     * @return 以EMU为单位的长度值
     */
    public int computeLengthInEMU(String length, String maxLength, int naturalEMU, int parentEMU) {
        int emu = naturalEMU;

        if (length.length() > 0) {
            CSSLength cssLength = CSSLength.of(length);
            if (cssLength.isValid()) {
                emu = computeLengthInEMU(cssLength, naturalEMU, parentEMU);
            }
        }

        if (maxLength.length() > 0) {
            CSSLength cssLength = CSSLength.of(maxLength);
            if (cssLength.isValid()) {
                int maxEMU = computeLengthInEMU(cssLength, naturalEMU, parentEMU);
                emu = Math.min(maxEMU, emu);
            }
        }

        return Math.min(emu, parentEMU);
    }

    /**
     * 考虑约束计算长度，以EMU为单位
     *
     * @param cssLength 长度声明
     * @param naturalEMU 原始长度
     * @param parentEMU 父容器长度
     * @return 以EMU为单位的长度值
     */
    public int computeLengthInEMU(CSSLength cssLength, int naturalEMU, int parentEMU) {
        int length;
        if (cssLength.getUnit() == CSSLengthUnit.PERCENT) {
            if (parentEMU != Integer.MAX_VALUE) {
                length = (int) (parentEMU * cssLength.getValue() * cssLength.getUnit().absoluteFactor());
            } else {
                length = naturalEMU;
            }
        } else {
            length = lengthToEMU(cssLength);
        }
        return length;
    }

    /**
     * 渲染文本
     *
     * @param text 文本
     */
    public void renderText(String text) {
        String whiteSpace = getPropertyValue(HtmlConstants.CSS_WHITE_SPACE);
        WhiteSpaceRule rule = WhiteSpaceRule.of(whiteSpace, WhiteSpaceRule.NORMAL);

        StringBuilder sb = StringUtil.borrowBuilder();
        boolean mergeWhitespace = false;

        // https://en.wikipedia.org/wiki/List_of_XML_and_HTML_character_entity_references
        int len = text.length();
        int c;

        if (!rule.isKeepTrailingSpace()) {
            boolean reachedLastNonWhite = false;
            for (int i = len - 1; i >= 0; i -= Character.charCount(c)) {
                c = text.codePointAt(i);
                switch (c) {
                    case ' ':
                    case '\r':
                    case '\n':
                    case '\t':
                    case 173: // soft hyphen
                    case 8203: // zero width space
                    case 8204: // zero width non-joiner
                    case 8205: // zero width joiner
                    case 8206: // lrm
                    case 8207: // rlm
                    case 8288: // word joiner
                    case 8289: // apply function
                    case 8290: // invisible times
                    case 8291: // invisible separator
                        len = i;
                        break;
                    default:
                        if (Character.getType(c) == 16) {
                            len = i;
                        } else {
                            reachedLastNonWhite = true;
                        }
                        break;
                }
                if (reachedLastNonWhite) {
                    break;
                }
            }
        }
        if (len == 0) {
            return;
        }
        boolean endTrimmed = len < text.length();
        CTR ctr = newRun();

        for (int i = 0; i < len; i += Character.charCount(c)) {
            c = text.codePointAt(i);
            switch (c) {
                case '\r':
                    if (i + 1 < len && text.codePointAt(i + 1) == '\n') {
                        continue;
                    }
                    if (rule.isKeepLineBreak()) {
                        addText(ctr, sb, false);
                        ctr.addNewCr();
                    } else {
                        mergeWhitespace = true;
                    }
                    break;
                case '\n':
                    if (rule.isKeepLineBreak()) {
                        addText(ctr, sb, false);
                        ctr.addNewBr();
                    } else {
                        mergeWhitespace = true;
                    }
                    break;
                case ' ':
                    if (rule.isKeepSpaceAndTab()) {
                        sb.appendCodePoint(c);
                    } else {
                        mergeWhitespace = true;
                    }
                    break;
                case '\t':
                    if (rule.isKeepSpaceAndTab()) {
                        addText(ctr, sb, false);
                        ctr.addNewTab();
                    } else {
                        mergeWhitespace = true;
                    }
                    break;
                case 160: // nbsp
                case 8192: // enquad
                case 8193: // emquad
                case 8194: // ensp
                case 8195: // emsp
                case 8196: // emsp13
                case 8197: // emsp14
                case 8199: // numsp
                case 8200: // puncsp
                case 8201: // thinsp
                case 8202: // hairsp
                case 8239: // narrow space
                case 8287: // medium space
                    if (mergeWhitespace) {
                        sb.append(' ');
                        mergeWhitespace = false;
                    }
                    sb.append(' ');
                    break;
                case 173: // soft hyphen
                case 8203: // zero width space
                case 8204: // zero width non-joiner
                case 8205: // zero width joiner
                case 8206: // lrm
                case 8207: // rlm
                case 8288: // word joiner
                case 8289: // apply function
                case 8290: // invisible times
                case 8291: // invisible separator
                    continue;
                default:
                    if (Character.getType(c) == 16) {
                        continue;
                    }
                    if (mergeWhitespace) {
                        if (sb.length() > 0) {
                            sb.append(' ');
                        }
                        mergeWhitespace = false;
                    }
                    sb.appendCodePoint(c);
                    if (previousText != null && previousText.isEndTrimmed()) {
                        CTText previous = previousText.getText();
                        previous.setStringValue(previous.getStringValue() + ' ');
                        previous.setSpace(SpaceAttribute.Space.PRESERVE);
                        previousText = null;
                    }
                    break;
            }
        }

        addText(ctr, sb, endTrimmed);
        StringUtil.releaseBuilder(sb);

        // 应用样式
        applyTextStyle(ctr);

        if (!(currentRun instanceof XWPFHyperlinkRun)) {
            currentRun = null;
        }
    }

    private void addText(CTR ctr, StringBuilder sb, boolean endTrimmed) {
        if (sb.length() > 0) {
            CTText ctText = ctr.addNewT();
            String text = sb.toString();
            ctText.setStringValue(text);
            if (text.charAt(0) == ' ' || text.charAt(sb.length() - 1) == ' ') {
                ctText.setSpace(SpaceAttribute.Space.PRESERVE);
            }
            sb.delete(0, sb.length());
            previousText = new TextWrapper(ctText, endTrimmed);
        }
    }

    /**
     * 应用文本样式
     *
     * @param ctr CTR
     */
    private void applyTextStyle(CTR ctr) {
        CTRPr rPr = RenderUtils.getRPr(ctr);

        // 字体，如果声明了全局字体则忽略样式声明
        String fontFamily = StringUtils.isBlank(globalFont) ? getPropertyValue(HtmlConstants.CSS_FONT_FAMILY) : globalFont;
        if (StringUtils.isNotBlank(fontFamily)) {
            CTFonts ctFonts = rPr.addNewRFonts();
            // ASCII
            ctFonts.setAscii(fontFamily);
            // High ANSI
            ctFonts.setHAnsi(fontFamily);
            // Complex Script
            ctFonts.setCs(fontFamily);
            // East Asian
            ctFonts.setEastAsia(fontFamily);
        }

        // 字号
        if (globalFontSize == null) {
            String fontSize = getPropertyValue(HtmlConstants.CSS_FONT_SIZE);
            if (StringUtils.isNotBlank(fontSize)) {
                int sz = getInheritedFontSizeInHalfPoints();
                rPr.addNewSz().setVal(BigInteger.valueOf(sz));
            }
        } else {
            // 如果定义了全局字号则忽略样式声明
            rPr.addNewSz().setVal(globalFontSize);
        }

        // 加粗
        String fontWeight = getPropertyValue(HtmlConstants.CSS_FONT_WEIGHT);
        if (fontWeight.contains(HtmlConstants.BOLD)) {
            rPr.addNewB();
        } else if (NumberUtils.isParsable(fontWeight) && Float.parseFloat(fontWeight) > 500) {
            rPr.addNewB();
        }

        // 斜体
        String fontStyle = getPropertyValue(HtmlConstants.CSS_FONT_STYLE);
        if (HtmlConstants.ITALIC.equals(fontStyle) || HtmlConstants.OBLIQUE.equals(fontStyle)) {
            rPr.addNewI();
        }

        // 颜色
        String color = getPropertyValue(HtmlConstants.CSS_COLOR);
        if (StringUtils.isNotBlank(color)) {
            String hex = Colors.fromStyle(color);
            RenderUtils.getColor(rPr).setVal(hex);
        }

        String caps = getPropertyValue(HtmlConstants.CSS_FONT_VARIANT_CAPS);
        if (HtmlConstants.SMALL_CAPS.equals(caps)) {
            rPr.addNewSmallCaps();
        }

        // 中划线/下划线
        String textDecoration = getPropertyValue(HtmlConstants.CSS_TEXT_DECORATION_LINE);
        if (HtmlConstants.NONE.equals(textDecoration)) {
            RenderUtils.getUnderline(rPr).setVal(STUnderline.NONE);
        } else {
            if (StringUtils.contains(textDecoration, HtmlConstants.LINE_THROUGH)) {
                rPr.addNewStrike();
            }
            if (StringUtils.contains(textDecoration, HtmlConstants.UNDERLINE)) {
                CTUnderline ctUnderline = RenderUtils.getUnderline(rPr);
                String textDecorationStyle = getPropertyValue(HtmlConstants.CSS_TEXT_DECORATION_STYLE);
                ctUnderline.setVal(RenderUtils.underline(textDecorationStyle));
                String textDecorationColor = getPropertyValue(HtmlConstants.CSS_TEXT_DECORATION_COLOR);
                if (StringUtils.isNotBlank(textDecorationColor)) {
                    String hex = Colors.fromStyle(textDecorationColor);
                    ctUnderline.setColor(hex);
                }
            }
        }

        // 上下标
        String verticalAlign = getPropertyValue(HtmlConstants.CSS_VERTICAL_ALIGN);
        if (HtmlConstants.SUPER.equals(verticalAlign)) {
            rPr.addNewVertAlign().setVal(STVerticalAlignRun.SUPERSCRIPT);
        } else if (HtmlConstants.SUB.equals(verticalAlign)) {
            rPr.addNewVertAlign().setVal(STVerticalAlignRun.SUBSCRIPT);
        }

        // FIXME 段落边框与行内边框分离，行内只有全边框，段落分四边

        // 背景色
        String backgroundColor = getPropertyValue(HtmlConstants.CSS_BACKGROUND_COLOR, true);
        if (StringUtils.isNotBlank(backgroundColor)) {
            String hex = Colors.fromStyle(backgroundColor, null);
            if (hex != null) {
                CTShd ctShd = rPr.addNewShd();
                ctShd.setFill(hex);
                ctShd.setVal(STShd.CLEAR);
            }
        }

        // 可见性
        String visibility = getPropertyValue(HtmlConstants.CSS_VISIBILITY);
        if (HtmlConstants.HIDDEN.equals(visibility) || HtmlConstants.COLLAPSE.equals(visibility)) {
            rPr.addNewVanish();
        }
    }

    /**
     * 渲染图片
     *
     * @param pictureData 图片数据流
     * @param pictureType 图片类型
     * @param filename 文件名
     * @param width 宽度
     * @param height 高度
     * @param svgData SVG数据
     */
    public void renderPicture(InputStream pictureData, int pictureType, String filename, int width, int height, byte[] svgData)
            throws IOException, InvalidFormatException {
        CTR ctr = newRun();

        XWPFPicture xwpfPicture = currentRun.addPicture(pictureData, pictureType, filename, width, height);
        CTR r = currentRun.getCTR();

        boolean isSvg = svgData != null;
        if (isSvg) {
            attachSvgData(xwpfPicture, svgData);
        }

        CSSStyleDeclarationImpl styleDeclaration = currentElementStyle();
        String cssFloat = styleDeclaration.getPropertyValue(HtmlConstants.CSS_FLOAT);
        boolean floatLeft = HtmlConstants.LEFT.equals(cssFloat);
        boolean floatRight = !floatLeft && HtmlConstants.RIGHT.equals(cssFloat);
        // vertical-align seems not working
        boolean floated = floatLeft || floatRight;

        CTDrawing drawing = null;
        if (r != ctr) {
            int lastDrawingIndex = r.sizeOfDrawingArray() - 1;
            drawing = r.getDrawingArray(lastDrawingIndex);
            ctr.setDrawingArray(new CTDrawing[]{drawing});
            r.removeDrawing(lastDrawingIndex);
            drawing = ctr.getDrawingArray(0);
        } else if (isSvg || floated) {
            drawing = ctr.getDrawingArray(ctr.sizeOfDrawingArray() - 1);
        }

        if (drawing != null && drawing.sizeOfInlineArray() > 0) {
            if (floated) {
                CTAnchor ctAnchor = RenderUtils.inlineToAnchor(drawing);

                CTPosH ctPosH = ctAnchor.addNewPositionH();
                ctPosH.setRelativeFrom(STRelFromH.MARGIN);
                ctPosH.setAlign(floatRight ? STAlignH.RIGHT : STAlignH.LEFT);

                CTPosV ctPosV = ctAnchor.addNewPositionV();
                ctPosV.setRelativeFrom(STRelFromV.PARAGRAPH);
                ctPosV.setAlign(STAlignV.TOP);

                if (isSvg) {
                    CTNonVisualGraphicFrameProperties properties = ctAnchor.addNewCNvGraphicFramePr();
                    CTGraphicalObjectFrameLocking frameLocking = properties.addNewGraphicFrameLocks();
                    frameLocking.setNoChangeAspect(true);
                }
            } else if (isSvg) {
                CTInline ctInline = drawing.getInlineArray(0);
                CTNonVisualGraphicFrameProperties properties = ctInline.isSetCNvGraphicFramePr()
                        ? ctInline.getCNvGraphicFramePr() : ctInline.addNewCNvGraphicFramePr();
                CTGraphicalObjectFrameLocking frameLocking = properties.isSetGraphicFrameLocks()
                        ? properties.getGraphicFrameLocks() : properties.addNewGraphicFrameLocks();
                frameLocking.setNoChangeAspect(true);
            }
        }
    }

    /**
     * 附加SVG数据
     *
     * @param xwpfPicture 图片
     * @param svgData SVG数据
     * @throws InvalidFormatException 非法格式
     */
    private void attachSvgData(XWPFPicture xwpfPicture, byte[] svgData) throws InvalidFormatException {
        CTPicture ctPicture = xwpfPicture.getCTPicture();
        String svgRelId = getXWPFDocument().addPictureData(svgData, SVGPictureData.PICTURE_TYPE_SVG);
        CTBlip blip = ctPicture.getBlipFill().getBlip();
        if (blip != null) {
            CTOfficeArtExtensionList extList = blip.isSetExtLst() ? blip.getExtLst() : blip.addNewExtLst();
            CTOfficeArtExtension svgBitmap = extList.addNewExt();
            svgBitmap.setUri(SVGRelation.SVG_URI);
            XmlCursor cur = svgBitmap.newCursor();
            cur.toEndToken();
            cur.beginElement(SVGRelation.SVG_QNAME);
            cur.insertNamespace(SVGRelation.SVG_PREFIX, SVGRelation.MS_SVG_NS);
            cur.insertAttributeWithValue(SVGRelation.EMBED_TAG, svgRelId);
            cur.dispose();
        }
    }

    /**
     * 将长度换算为EMU
     *
     * @param length 长度
     * @return EMU
     */
    public int lengthToEMU(CSSLength length) {
        if (!length.isValid()) {
            throw new UnsupportedOperationException("Invalid CSS length");
        }
        if (!length.getUnit().isRelative()) {
            return length.toEMU();
        }
        double emu;
        switch (length.getUnit()) {
            case REM:
                emu = length.unitValue() * getDefaultFontSize().toEMU();
                break;
            case EM:
                emu = length.unitValue() * getInheritedFontSizeInHalfPoints() * Units.EMU_PER_POINT / 2;
                break;
            case VW:
                emu = length.unitValue() * getPageWidth().toEMU();
                break;
            case VH:
                emu = length.unitValue() * getPageHeight().toEMU();
                break;
            case VMIN:
                emu = length.unitValue() * Math.min(getPageWidth().toEMU(), getPageHeight().toEMU());
                break;
            case VMAX:
                emu = length.unitValue() * Math.max(getPageWidth().toEMU(), getPageHeight().toEMU());
                break;
            // Unable to determine the use of width or height as a relative length for percent unit
            default:
                throw new UnsupportedOperationException("Can not convert to EMU with length: " + length);
        }
        return (int) Math.rint(emu);
    }

    public NumberingContext getNumberingContext() {
        return numberingContext;
    }

    public CSSLength getPageWidth() {
        return pageWidth;
    }

    public CSSLength getPageHeight() {
        return pageHeight;
    }

    public CSSLength getMarginTop() {
        return marginTop;
    }

    public CSSLength getMarginRight() {
        return marginRight;
    }

    public CSSLength getMarginBottom() {
        return marginBottom;
    }

    public CSSLength getMarginLeft() {
        return marginLeft;
    }

    public int getAvailablePageWidth() {
        return availablePageWidth;
    }

    public int getAvailablePageHeight() {
        return availablePageHeight;
    }

    public XWPFRun getCurrentRun() {
        return currentRun;
    }

    public String getGlobalFont() {
        return globalFont;
    }

    public BigInteger getGlobalFontSize() {
        return globalFontSize;
    }

    public boolean isShowDefaultTableBorderInTableCell() {
        return showDefaultTableBorderInTableCell;
    }

    public void setShowDefaultTableBorderInTableCell(boolean showDefaultTableBorderInTableCell) {
        this.showDefaultTableBorderInTableCell = showDefaultTableBorderInTableCell;
    }

    public void setGlobalFont(String globalFont) {
        this.globalFont = globalFont;
    }

    public void setGlobalFontSize(BigInteger globalFontSize) {
        this.globalFontSize = globalFontSize;
    }

    public boolean isBlocked() {
        return blockLevel > 0;
    }

    public void incrementBlockLevel() {
        blockLevel++;
    }

    public void decrementBlockLevel() {
        blockLevel--;
    }

    public void renderDocument(Document document) {
        Element body = document.body();
        Element html = body.parent();
        if (html.hasAttr(HtmlConstants.ATTR_STYLE)) {
            pushInlineStyle(getCssStyleDeclaration(html), html.isBlock());
        }
        if (body.hasAttr(HtmlConstants.ATTR_STYLE)) {
            pushInlineStyle(getCssStyleDeclaration(body), body.isBlock());
        }
        for (Node node : body.childNodes()) {
            renderNode(node);
        }
    }

    public void renderNode(Node node) {
        boolean isElement = node instanceof Element;

        if (isElement) {
            Element element = ((Element) node);
            renderElement(element);
        } else if (node instanceof TextNode && !"tr".equals(node.parent().nodeName())) {
            renderText(((TextNode) node).getWholeText());
        }
    }

    public void renderElement(Element element) {
        if (log.isDebugEnabled()) {
            log.info("Start rendering html tag: <{}{}>", element.normalName(), element.attributes());
        }
        if (element.tag().isFormListed() || element.tag().isFormSubmittable()) {
            return;
        }

        CSSStyleDeclarationImpl cssStyleDeclaration = getCssStyleDeclaration(element);
        String display = cssStyleDeclaration.getPropertyValue(HtmlConstants.CSS_DISPLAY);
        if (HtmlConstants.NONE.equalsIgnoreCase(display)) {
            return;
        }
        pushInlineStyle(cssStyleDeclaration, element.isBlock());

        ElementRenderer elementRenderer = rendererProvider.get(element.normalName());
        boolean blocked = false;

        if (renderAsBlock(element, elementRenderer)) {
            if (element.childNodeSize() == 0 && !HtmlConstants.KEEP_EMPTY_TAGS.contains(element.normalName())) {
                popInlineStyle();
                return;
            }
            if (!isBlocked()) {
                // 复制段落中占位符之前的部分内容
                moveContentToNewPrevParagraph();
            }
            incrementBlockLevel();
            blocked = true;

            IBody container = getContainer();
            boolean isTableTag = HtmlConstants.TAG_TABLE.equals(element.normalName());

            adjustCursor(container, isTableTag);

            if (isTableTag) {
                globalCursor.push();
                XWPFTable xwpfTable = container.insertNewTbl(globalCursor);
                globalCursor.pop();
                if (dedupeParagraph != null && !numberingContext.contains(dedupeParagraph)) {
                    if (!dedupeParagraph.equals(getRun().getParent())) {
                        tryRemoveParagraph(container);
                    }
                    unmarkDedupe();
                }
                // 新增时会自动创建一行一列，会影响自定义的表格渲染逻辑，故删除
                xwpfTable.removeRow(0);

                if (container.getPartType() == BodyType.TABLECELL && isShowDefaultTableBorderInTableCell()) {
                    CTTbl ctTbl = xwpfTable.getCTTbl();
                    CTTblPr tblPr = RenderUtils.getTblPr(ctTbl);
                    CTTblBorders tblBorders = RenderUtils.getTblBorders(tblPr);
                    tblBorders.addNewTop().setVal(STBorder.SINGLE);
                    tblBorders.addNewLeft().setVal(STBorder.SINGLE);
                    tblBorders.addNewBottom().setVal(STBorder.SINGLE);
                    tblBorders.addNewRight().setVal(STBorder.SINGLE);
                    tblBorders.addNewInsideH().setVal(STBorder.SINGLE);
                    tblBorders.addNewInsideV().setVal(STBorder.SINGLE);
                }

                RenderUtils.tableStyle(this, xwpfTable, cssStyleDeclaration);
            } else if (shouldNewParagraph(element)) {
                globalCursor.push();
                XWPFParagraph xwpfParagraph = newParagraph(container, globalCursor);
                globalCursor.pop();
                if (xwpfParagraph == null) {
                    log.warn("Can not add new paragraph for element: {}, attributes: {}", element.tagName(), element.attributes().html());
                }

                RenderUtils.paragraphStyle(this, xwpfParagraph, cssStyleDeclaration);
            } else {
                RenderUtils.paragraphStyle(this, dedupeParagraph, cssStyleDeclaration);
            }
        }

        if (elementRenderer != null) {
            if (!elementRenderer.renderStart(element, this)) {
                renderElementEnd(element, this, elementRenderer, blocked);
                return;
            }
        }

        for (Node child : element.childNodes()) {
            renderNode(child);
        }

        renderElementEnd(element, this, elementRenderer, blocked);
    }

    private void tryRemoveParagraph(IBody container) {
        System.out.println("tryRemoveParagraph");
        List<XWPFRun> runs = dedupeParagraph.getRuns();
        int size = runs.size();
        if (size == 0) {
            this.removeParagraph(container, dedupeParagraph);
            return;
        }
        int index = runs.size() - 1;
        XWPFRun run = runs.get(index);
        String text = run.text();
        if ((text.isEmpty() || "\n".equals(text)) && run.getEmbeddedPictures().isEmpty()) {
            CTP ctp = dedupeParagraph.getCTP();
            if (size == 1 && ctp.sizeOfOMathArray() == 0 && ctp.sizeOfOMathParaArray() == 0) {
                this.removeParagraph(container, dedupeParagraph);
            } else {
                dedupeParagraph.removeRun(index);
            }
        }
    }

    private void removeParagraph(IBody container, XWPFParagraph paragraph) {
        switch (container.getPartType()) {
            case CONTENTCONTROL:
                break;
            case DOCUMENT:
                XWPFDocument xwpfDocument = (XWPFDocument) container;
                int posOfParagraph = xwpfDocument.getPosOfParagraph(paragraph);
                xwpfDocument.removeBodyElement(posOfParagraph);
                break;
            case HEADER:
                XWPFHeader xwpfHeader = (XWPFHeader) container;
                xwpfHeader.removeParagraph(paragraph);
                break;
            case FOOTER:
                XWPFFooter xwpfFooter = (XWPFFooter) container;
                xwpfFooter.removeParagraph(paragraph);
                break;
            case FOOTNOTE:
                XWPFFootnote xwpfFootnote = (XWPFFootnote) container;
                xwpfFootnote.getParagraphs().remove(paragraph);
                break;
            case TABLECELL:
                XWPFTableCell xwpfTableCell = (XWPFTableCell) container;
                xwpfTableCell.removeParagraph(xwpfTableCell.getParagraphs().indexOf(paragraph));
                break;
        }
    }

    /**
     * HTML元素是否按照块状进行渲染
     *
     * @param element HTML元素
     * @return 是否按照块状进行渲染
     */
    public boolean renderAsBlock(Element element) {
        return renderAsBlock(element, rendererProvider.get(element.normalName()));
    }

    /**
     * HTML元素是否按照块状进行渲染
     *
     * @param element HTML元素
     * @param elementRenderer 元素渲染器
     * @return 是否按照块状进行渲染
     */
    private boolean renderAsBlock(Element element, ElementRenderer elementRenderer) {
        return element.isBlock() && (elementRenderer == null || elementRenderer.renderAsBlock());
    }

    private boolean shouldNewParagraph(Element element) {
        if (dedupeParagraph == null) {
            return true;
            // return !HtmlConstants.TAG_HR.equals(element.normalName());
        }
        boolean newParagraph = false;
        XmlCursor xmlCursor = dedupeParagraph.getCTP().newCursor();
        if (xmlCursor.toPrevSibling()) {
            if (XmlUtils.P_QNAME.equals(xmlCursor.getName())) {
                if (xmlCursor.toLastChild()) {
                    if (XmlUtils.R_QNAME.equals(xmlCursor.getName())) {
                        xmlCursor.push();
                        if (xmlCursor.toFirstChild() && XmlUtils.BR_QNAME.equals(xmlCursor.getName()) && !xmlCursor.toNextSibling()) {
                            xmlCursor.pop();
                            xmlCursor.removeXml();
                            unmarkDedupe();
                            newParagraph = true;
                        }
                    }
                }
            }
        }
        xmlCursor.dispose();
        return newParagraph;
    }

    private void adjustCursor(IBody container, boolean isTableTag) {
        if (XmlUtils.R_QNAME.equals(globalCursor.getName())) {
            globalCursor.push();
            globalCursor.toParent();
        }
        globalCursor.push();
        // 如果是表格，检查当前word容器的前一个兄弟元素是否为表格，是则插入一个段落，防止表格粘连在一起
        if (isTableTag && globalCursor.toPrevSibling()) {
            if (XmlUtils.TBL_QNAME.equals(globalCursor.getName())) {
                // pop() is safer than toNextSibling()
                globalCursor.pop();
                globalCursor.push();
                XWPFParagraph paragraph = newParagraph(container, globalCursor);
                unmarkDedupe();
                RenderUtils.paragraphStyle(this, paragraph, CSSStyleUtils.EMPTY_STYLE);
            }
        }
        globalCursor.pop();
    }

    private void renderElementEnd(Element element, HtmlRenderContext context, ElementRenderer elementRenderer, boolean blocked) {
        if (elementRenderer != null) {
            elementRenderer.renderEnd(element, context);
        }
        context.popInlineStyle();
        if (blocked) {
            context.decrementBlockLevel();
        }
    }

    private void moveContentToNewPrevParagraph() {
        CTR ctr = getRun().getCTR();
        XmlCursor rCursor = ctr.newCursor();
        boolean hasPrevSibling = false;
        while (rCursor.toPrevSibling()) {
            XmlObject object = rCursor.getObject();
            if (object instanceof CTMarkupRange) {
                continue;
            }
            if (!XmlUtils.PPR_QNAME.equals(rCursor.getName())) {
                hasPrevSibling = true;
                break;
            }
        }
        if (!hasPrevSibling) {
            rCursor.dispose();
            return;
        }
        rCursor.toParent();
        rCursor.push();
        CTP ctp = ((CTP) rCursor.getObject());
        XWPFParagraph paragraph = getContainer().getParagraph(ctp);
        XWPFParagraph newParagraph = getContainer().insertNewParagraph(rCursor);
        XmlCursor pCursor = newParagraph.getCTP().newCursor();
        pCursor.toEndToken();
        rCursor.pop();
        rCursor.toFirstChild();
        while (!ctr.equals(rCursor.getObject())) {
            if (XmlUtils.PPR_QNAME.equals(rCursor.getName())) {
                rCursor.copyXml(pCursor);
                rCursor.toNextSibling();
            } else if (XmlUtils.BOOKMARK_START_QNAME.equals(rCursor.getName()) || XmlUtils.BOOKMARK_END_QNAME.equals(rCursor.getName())) {
                rCursor.toNextSibling();
            } else {
                // moveXml附带了toNextSibling的效果
                rCursor.moveXml(pCursor);
            }
        }
        rCursor.dispose();
        pCursor.dispose();

        XWPFParagraphRuns runs = new XWPFParagraphRuns(paragraph);

        for (int i = runs.runCount() - ctp.getRList().size() - 1; i >= 0; i--) {
            runs.remove(i);
        }
    }

    public CSSStyleDeclarationImpl getCssStyleDeclaration(Element element) {
        String style = element.attr(HtmlConstants.ATTR_STYLE);
        CSSStyleDeclarationImpl cssStyleDeclaration = CSSStyleUtils.parse(style);
        CSSStyleUtils.split(cssStyleDeclaration);
        return cssStyleDeclaration;
    }

    /**
     * 保存当前指针位置并移动到目标指针位置
     *
     * @param targetCursor 目标指针
     */
    public void pushCursor(XmlCursor targetCursor) {
        globalCursor.push();
        globalCursor.toCursor(targetCursor);
    }

    /**
     * 返回之前保存的指针位置
     *
     * @return 是否返回成功
     */
    public boolean popCursor() {
        return globalCursor.pop();
    }

    /**
     * @return 指针当前指向的对象
     */
    public XmlObject currentCursorObject() {
        return globalCursor.getObject();
    }

    /**
     * 标记段落以防止块状元素嵌套产生多余的空段落
     *
     * @param paragraph 段落
     */
    public void markDedupe(XWPFParagraph paragraph) {
        dedupeParagraph = paragraph;
    }

    /**
     * 取消段落防重标记
     */
    public void unmarkDedupe() {
        dedupeParagraph = null;
    }

    /**
     * 文本封装类，用于空白字符折叠处理
     */
    private static class TextWrapper {
        private final CTText text;
        private final boolean endTrimmed;

        public TextWrapper(CTText text, boolean endTrimmed) {
            this.text = text;
            this.endTrimmed = endTrimmed;
        }

        public CTText getText() {
            return text;
        }

        public boolean isEndTrimmed() {
            return endTrimmed;
        }
    }
}
