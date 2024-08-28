/*
 * Copyright 2016 - 2022 Draco, https://github.com/draco1023
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

import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import org.apache.commons.lang3.StringUtils;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.jsoup.nodes.Element;

/**
 * 标记标签渲染器
 *
 * @author Draco
 * @since 2022-06-11
 */
public class MarkRenderer implements ElementRenderer {
    private static final String[] TAGS = {HtmlConstants.TAG_MARK};

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        CSSStyleDeclarationImpl cssStyleDeclaration = context.currentElementStyle();

        if (StringUtils.isEmpty(cssStyleDeclaration.getBackgroundColor())) {
            cssStyleDeclaration.setBackgroundColor("yellow");
        }
        return true;
    }

    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    @Override
    public boolean renderAsBlock() {
        return false;
    }
}
