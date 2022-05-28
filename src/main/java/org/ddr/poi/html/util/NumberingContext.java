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

package org.ddr.poi.html.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHint;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMultiLevelType;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

/**
 * 列表上下文
 *
 * @author Draco
 * @since 2021-02-19
 */
public class NumberingContext {
    /**
     * 每级缩进
     */
    private int indent = 360;
    private final XWPFDocument document;
    private int nextAbstractNumberId;
    private int nextNumberingLevel;

    private List<ListStyleType> numberFormats;
    private TreeMap<String, BigInteger> numberIdMap = new TreeMap<>(Collections.reverseOrder());
    private List<XWPFParagraph> numberingParagraphs;

    public NumberingContext(XWPFDocument document) {
        this.document = document;
    }

    /**
     * 开始新的列表
     *
     * @param format 列表符号类型
     */
    public void startLevel(ListStyleType format) {
        int level = nextNumberingLevel++;
        if (level == 0) {
            numberingParagraphs = new ArrayList<>(8);
            numberFormats = new ArrayList<>(4);
        }
        numberFormats.add(format);
    }

    /**
     * 结束当前列表
     */
    public void endLevel() {
        nextNumberingLevel--;

        String key = getFormatKey();
        BigInteger numberId = getNumberId(key);

        BigInteger currentLevel = BigInteger.valueOf(nextNumberingLevel);
        for (int i = numberingParagraphs.size() - 1; i >= 0; i--) {
            XWPFParagraph paragraph = numberingParagraphs.get(i);
            if (currentLevel.equals(paragraph.getNumIlvl())) {
                paragraph.setNumID(numberId);
                numberingParagraphs.remove(i);
            } else {
                break;
            }
        }
        if (!numberFormats.isEmpty()) {
            numberFormats.remove(numberFormats.size() - 1);
        }

        if (nextNumberingLevel == 0) {
            numberingParagraphs = null;
            numberFormats = null;
        }
    }

    /**
     * 新增段落
     *
     * @param paragraph 段落
     */
    public void add(XWPFParagraph paragraph) {
        if (numberingParagraphs == null) {
            throw new IllegalStateException("Call startLevel method first");
        }
        paragraph.setNumILvl(BigInteger.valueOf(nextNumberingLevel - 1));
        numberingParagraphs.add(paragraph);
    }

    /**
     * Setter for indentation.
     *
     * @param indentation New indentation value
     */
    public void setIndentation(int indentation) {
        this.indent = indentation;
    }

    /**
     * 获取列表ID
     *
     * @param key Key
     * @return 列表ID
     */
    private BigInteger getNumberId(String key) {
        BigInteger numberId = null;
        for (Map.Entry<String, BigInteger> entry : numberIdMap.entrySet()) {
            if (entry.getKey().startsWith(key)) {
                numberId = entry.getValue();
                break;
            }
        }
        if (numberId == null) {
            XWPFNumbering numbering = document.createNumbering();
            while (true) {
                BigInteger abstractNumberId = BigInteger.valueOf(nextAbstractNumberId++);
                XWPFAbstractNum abstractNum = numbering.getAbstractNum(abstractNumberId);
                if (abstractNum == null) {
                    CTAbstractNum ctAbstractNum = CTAbstractNum.Factory.newInstance();
                    ctAbstractNum.setAbstractNumId(abstractNumberId);
                    ctAbstractNum.addNewMultiLevelType().setVal(STMultiLevelType.HYBRID_MULTILEVEL);

                    for (int i = 0; i < numberFormats.size(); i++) {
                        ListStyleType listStyleType = numberFormats.get(i);
                        CTLvl cTLvl = ctAbstractNum.addNewLvl();
                        CTInd ind = cTLvl.addNewPPr().addNewInd();
                        ind.setLeft(BigInteger.valueOf((long) indent * i));

                        cTLvl.addNewNumFmt().setVal(listStyleType.getFormat());
                        cTLvl.addNewLvlText().setVal(getLevelText(listStyleType, i));
                        cTLvl.addNewStart().setVal(BigInteger.ONE);
                        cTLvl.setIlvl(BigInteger.valueOf(i));
                        cTLvl.addNewLvlJc().setVal(STJc.LEFT);
                        if (StringUtils.isNotBlank(listStyleType.getFont())) {
                            CTFonts ctFonts = cTLvl.addNewRPr().addNewRFonts();
                            ctFonts.setAscii(listStyleType.getFont());
                            ctFonts.setHAnsi(listStyleType.getFont());
                            ctFonts.setHint(STHint.DEFAULT);
                        }
                    }

                    numbering.addAbstractNum(new XWPFAbstractNum(ctAbstractNum, numbering));
                    numberId = numbering.addNum(abstractNumberId);

                    numberIdMap.put(key, numberId);
                    break;
                }
            }
        }
        return numberId;
    }

    private String getLevelText(ListStyleType listStyleType, int i) {
        if (listStyleType.getText() == null) {
            return "";
        }
        if (listStyleType.getText().length() == 0) {
            return getOrderedLevelText(i);
        }
        return listStyleType.getText();
    }

    /**
     * 获取有序列表项的序号格式
     *
     * @param i 索引，从0开始
     * @return 序号格式
     */
    private String getOrderedLevelText(int i) {
        return "%" + (i + 1) + ".";
    }

    private String getFormatKey() {
        StringBuilder sb = new StringBuilder();
        for (ListStyleType format : numberFormats) {
            sb.append(format.getName()).append(StringUtils.SPACE);
        }
        return sb.toString();
    }

}
