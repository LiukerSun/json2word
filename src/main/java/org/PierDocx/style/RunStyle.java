package org.PierDocx.style;

import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;

import java.io.Serializable;

public class RunStyle implements Serializable {
    private Boolean bold;
    private Boolean Capitalized;
    private int CharacterSpacing;// 字符间距
    private java.lang.String Color;
    private Boolean DoubleStrikethrough;
    private Boolean Embossed;// 着重
    private java.lang.String EmphasisMark;//着重符号
    private java.lang.String FontFamily;
    private String westernFontFamily; // western font
    private double FontSize;
    private Boolean Imprinted;
    private Boolean Italic;
    private int Kerning; // 字体调整字间距
    private java.lang.String Lang; // 设置段落关联语言
    private Boolean Shadow;
    private Boolean SmallCaps;
    private Boolean Strike; //删除线
    private Boolean StrikeThrough;
    private VerticalAlign Subscript; //上标/下标
    private java.lang.String TextHighlightColor;
    private UnderlinePatterns Underline;
    private java.lang.String UnderlineColor;
    private java.lang.String VerticalAlignment;
    private String vertAlign;
    private XWPFHighlightColor highlightColor;
    private UnderlinePatterns underlinePatterns;


    // setter
    public RunStyle setBold(Boolean bold) {
        this.bold = bold;
        return this;
    }

    public RunStyle setCapitalized(Boolean Capitalized) {
        this.Capitalized = Capitalized;
        return this;
    }

    public RunStyle setCharacterSpacing(int twips) {
        this.CharacterSpacing = twips;
        return this;
    }

    public RunStyle setColor(java.lang.String rgbStr) {
        this.Color = rgbStr;
        return this;
    }

    public RunStyle setDoubleStrikethrough(Boolean DoubleStrikethrough) {
        this.DoubleStrikethrough = DoubleStrikethrough;
        return this;
    }

    public RunStyle setEmbossed(Boolean Embossed) {
        this.Embossed = Embossed;
        return this;
    }

    public RunStyle setEmphasisMark(java.lang.String EmphasisMark) {
        this.EmphasisMark = EmphasisMark;
        return this;
    }

    public RunStyle setFontFamily(java.lang.String FontFamily) {
        this.FontFamily = FontFamily;
        return this;
    }

    public RunStyle setFontSize(int FontSize) {
        this.FontSize = FontSize;
        return this;
    }

    public RunStyle setFontSize(double FontSize) {
        this.FontSize = FontSize;
        return this;
    }

    public RunStyle setImprinted(Boolean Imprinted) {
        this.Imprinted = Imprinted;
        return this;
    }

    public RunStyle setItalic(Boolean Italic) {
        this.Italic = Italic;
        return this;
    }

    public RunStyle setKerning(int Kerning) {
        this.Kerning = Kerning;
        return this;
    }

    public RunStyle setLang(java.lang.String Lang) {
        this.Lang = Lang;
        return this;
    }

    public RunStyle setShadow(Boolean Shadow) {
        this.Shadow = Shadow;
        return this;
    }

    public RunStyle setSmallCaps(Boolean SmallCaps) {
        this.SmallCaps = SmallCaps;
        return this;
    }

    public RunStyle setStrike(Boolean Strike) {
        this.Strike = Strike;
        return this;
    }

    public RunStyle setStrikeThrough(Boolean StrikeThrough) {
        this.StrikeThrough = StrikeThrough;
        return this;
    }

    public RunStyle setSubscript(VerticalAlign valign) {
        this.Subscript = valign;
        return this;
    }

    public RunStyle setTextHighlightColor(java.lang.String colorName) {
        this.TextHighlightColor = colorName;
        return this;
    }

    public RunStyle setUnderline(UnderlinePatterns value) {
        this.Underline = value;
        return this;
    }

    public RunStyle setUnderlineColor(java.lang.String color) {
        this.UnderlineColor = color;
        return this;
    }

    public RunStyle VerticalAlignment(java.lang.String verticalAlignment) {
        this.VerticalAlignment = VerticalAlignment;
        return this;
    }

    public RunStyle setWesternFontFamily(String westernFontFamily) {
        this.westernFontFamily = westernFontFamily;
        return this;

    }

    public RunStyle setVertAlign(String vertAlign) {
        this.vertAlign = vertAlign;
        return this;
    }

    public RunStyle setHighlightColor(XWPFHighlightColor highlightColor) {
        this.highlightColor = highlightColor;
        return this;
    }

    public RunStyle setUnderlinePatterns(UnderlinePatterns underlinePatterns) {
        this.underlinePatterns = underlinePatterns;
        return this;
    }


    // getter
    public Boolean isBold() {
        return bold;
    }

    public Boolean isCapitalized() {
        return Capitalized;
    }

    public int getCharacterSpacing() {
        return CharacterSpacing;
    }

    public String getColor() {
        return Color;
    }

    public Boolean isDoubleStrikethrough() {
        return DoubleStrikethrough;
    }

    public Boolean isEmbossed() {
        return Embossed;
    }

    public String getEmphasisMark() {
        return EmphasisMark;
    }

    public String getFontFamily() {
        return FontFamily;
    }

    public double getFontSize() {
        return FontSize;
    }

    public Boolean isImprinted() {
        return Imprinted;
    }

    public Boolean isItalic() {
        return Italic;
    }

    public int getKerning() {
        return Kerning;
    }

    public String getLang() {
        return Lang;
    }

    public Boolean isShadow() {
        return Shadow;
    }

    public Boolean isSmallCaps() {
        return SmallCaps;
    }

    public Boolean isStrike() {
        return Strike;
    }

    public Boolean isStrikeThrough() {
        return StrikeThrough;
    }

    public VerticalAlign getSubscript() {
        return Subscript;
    }

    public String getTextHighlightColor() {
        return TextHighlightColor;
    }

    public UnderlinePatterns getUnderline() {
        return Underline;
    }

    public String getUnderlineColor() {
        return UnderlineColor;
    }

    public String getVerticalAlignment() {
        return VerticalAlignment;
    }

    public String getWesternFontFamily() {
        return westernFontFamily;
    }

    public String getVertAlign() {
        return vertAlign;
    }

    public XWPFHighlightColor getHighlightColor() {
        return highlightColor;
    }

    public UnderlinePatterns getUnderlinePatterns() {
        return underlinePatterns;
    }

}
