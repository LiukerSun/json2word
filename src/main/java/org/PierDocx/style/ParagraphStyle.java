package org.PierDocx.style;

import java.io.Serializable;

import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

public class ParagraphStyle implements Serializable {
    private String styleId;
    private ParagraphAlignment align; // 对齐
    private Double indentLeftChars; // 缩进-左侧
    private Double indentRightChars;// 缩进-右侧
    private Double indentHangingChars;// 缩进-悬挂
    private Double indentFirstLineChars;// 缩进-首行
    private BorderStyle leftBorder; // 边框-左
    private BorderStyle rightBorder;// 边框-右
    private BorderStyle topBorder;// 边框-上
    private BorderStyle bottomBorder;// 边框-下
    private XWPFShadingPattern shadingPattern; // 底纹样式
    private String backgroundColor; // 底纹颜色
    private Boolean widowControl; // 孤行控制
    private Boolean keepLines; // 段中不分页
    private Boolean keepNext; // 与下段同页
    private Boolean pageBreakBefore; // 段前分页
    private Boolean allowWordBreak; // 单词断行
    private Double spacingBeforeLines; // 间距-段前-行
    private Double spacingAfterLines; // 间距-断后-行
    private Double spacingBefore; // 间距-段前-磅
    private Double spacingAfter; // 间距-断后-磅
    private Double spacing;
    private LineSpacingRule spacingRule;
    private Style glyphStyle;
    private Long numId;
    private Long lvl;

    public ParagraphStyle setStyleId(String styleId) {
        this.styleId = styleId;
        return this;
    }

    public ParagraphStyle setAlign(ParagraphAlignment align) {
        this.align = align;
        return this;
    }

    public ParagraphStyle setIndentLeftChars(Double indentLeftChars) {
        this.indentLeftChars = indentLeftChars;
        return this;
    }

    public ParagraphStyle setIndentRightChars(Double indentRightChars) {
        this.indentRightChars = indentRightChars;
        return this;
    }

    public ParagraphStyle setIndentHangingChars(Double indentHangingChars) {
        this.indentHangingChars = indentHangingChars;
        return this;
    }

    public ParagraphStyle setIndentFirstLineChars(Double indentFirstLineChars) {
        this.indentFirstLineChars = indentFirstLineChars;
        return this;
    }

    public ParagraphStyle setLeftBorder(BorderStyle leftBorder) {
        this.leftBorder = leftBorder;
        return this;
    }

    public ParagraphStyle setRightBorder(BorderStyle rightBorder) {
        this.rightBorder = rightBorder;
        return this;
    }

    public ParagraphStyle setTopBorder(BorderStyle topBorder) {
        this.topBorder = topBorder;
        return this;
    }

    public ParagraphStyle setBottomBorder(BorderStyle bottomBorder) {
        this.bottomBorder = bottomBorder;
        return this;
    }

    public ParagraphStyle setShadingPattern(XWPFShadingPattern shadingPattern) {
        this.shadingPattern = shadingPattern;
        return this;
    }

    public ParagraphStyle setBackgroundColor(String backgroundColor) {
        this.backgroundColor = backgroundColor;
        return this;
    }

    public ParagraphStyle setWidowControl(Boolean widowControl) {
        this.widowControl = widowControl;
        return this;
    }

    public ParagraphStyle setKeepLines(Boolean keepLines) {
        this.keepLines = keepLines;
        return this;
    }

    public ParagraphStyle setKeepNext(Boolean keepNext) {
        this.keepNext = keepNext;
        return this;
    }

    public ParagraphStyle setPageBreakBefore(Boolean pageBreakBefore) {
        this.pageBreakBefore = pageBreakBefore;
        return this;
    }

    public ParagraphStyle setAllowWordBreak(Boolean allowWordBreak) {
        this.allowWordBreak = allowWordBreak;
        return this;
    }

    public ParagraphStyle setSpacingBeforeLines(Double spacingBeforeLines) {
        this.spacingBeforeLines = spacingBeforeLines;
        return this;
    }

    public ParagraphStyle setSpacingAfterLines(Double spacingAfterLines) {
        this.spacingAfterLines = spacingAfterLines;
        return this;
    }

    public ParagraphStyle setSpacingBefore(Double spacingBefore) {
        this.spacingBefore = spacingBefore;
        return this;
    }

    public ParagraphStyle setSpacingAfter(Double spacingAfter) {
        this.spacingAfter = spacingAfter;
        return this;
    }

    public ParagraphStyle setSpacing(Double spacing) {
        this.spacing = spacing;
        return this;
    }

    public ParagraphStyle setSpacingRule(LineSpacingRule spacingRule) {
        this.spacingRule = spacingRule;
        return this;
    }

    public ParagraphStyle setGlyphStyle(Style glyphStyle) {
        this.glyphStyle = glyphStyle;
        return this;
    }

    public ParagraphStyle setNumId(Long numId) {
        this.numId = numId;
        return this;
    }

    public ParagraphStyle setLvl(Long lvl) {
        this.lvl = lvl;
        return this;
    }


    // getter
    public String getStyleId() {
        return styleId;
    }

    public ParagraphAlignment getAlign() {
        return align;
    }

    public Double getIndentLeftChars() {
        return indentLeftChars;
    }

    public Double getIndentRightChars() {
        return indentRightChars;
    }

    public Double getIndentHangingChars() {
        return indentHangingChars;
    }

    public Double getIndentFirstLineChars() {
        return indentFirstLineChars;
    }

    public BorderStyle getLeftBorder() {
        return leftBorder;
    }

    public BorderStyle getRightBorder() {
        return rightBorder;
    }

    public BorderStyle getTopBorder() {
        return topBorder;
    }

    public BorderStyle getBottomBorder() {
        return bottomBorder;
    }

    public XWPFShadingPattern getShadingPattern() {
        return shadingPattern;
    }

    public String getBackgroundColor() {
        return backgroundColor;
    }

    public Boolean getWidowControl() {
        return widowControl;
    }

    public Boolean getKeepLines() {
        return keepLines;
    }

    public Boolean getKeepNext() {
        return keepNext;
    }

    public Boolean getPageBreakBefore() {
        return pageBreakBefore;
    }

    public Boolean getAllowWordBreak() {
        return allowWordBreak;
    }

    public Double getSpacingBeforeLines() {
        return spacingBeforeLines;
    }

    public Double getSpacingAfterLines() {
        return spacingAfterLines;
    }

    public Double getSpacingBefore() {
        return spacingBefore;
    }

    public Double getSpacingAfter() {
        return spacingAfter;
    }

    public Double getSpacing() {
        return spacing;
    }

    public LineSpacingRule getSpacingRule() {
        return spacingRule;
    }

    public Style getGlyphStyle() {
        return glyphStyle;
    }

    public Long getNumId() {
        return numId;
    }

    public Long getLvl() {
        return lvl;
    }


}
