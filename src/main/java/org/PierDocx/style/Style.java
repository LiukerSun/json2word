package org.PierDocx.style;

import java.io.Serial;
import java.io.Serializable;

import org.apache.poi.xwpf.usermodel.UnderlinePatterns;

public class Style implements Serializable {

    @Serial
    private static final long serialVersionUID = 1L;

    private String color;
    private String fontFamily; // east Asia font
    private String westernFontFamily; // western font
    private double fontSize;
    private Boolean isBold;
    private Boolean isItalic;
    private Boolean isStrike;
    private UnderlinePatterns underlinePatterns;
    private String underlineColor;
    private XWPFHighlightColor highlightColor;
    private int characterSpacing;
    private String vertAlign;


    public Style() {
    }

    public Style(String color) {
        this.color = color;
    }

    public Style(String fontFamily, double fontSize) {
        this.fontFamily = fontFamily;
        this.fontSize = fontSize;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public void setFontFamily(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public void setFontSize(double fontSize) {
        this.fontSize = fontSize;
    }

    public void setBold(Boolean isBold) {
        this.isBold = isBold;
    }

    public void setItalic(Boolean isItalic) {
        this.isItalic = isItalic;
    }

    public void setStrike(Boolean isStrike) {
        this.isStrike = isStrike;
    }

    public void setUnderlinePatterns(UnderlinePatterns underlinePatterns) {
        this.underlinePatterns = underlinePatterns;
    }

    public void setUnderlineColor(String underlineColor) {
        this.underlineColor = underlineColor;
    }

    public void setHighlightColor(XWPFHighlightColor highlightColor) {
        this.highlightColor = highlightColor;
    }

    public void setCharacterSpacing(int characterSpacing) {
        this.characterSpacing = characterSpacing;
    }

    public void setVertAlign(String vertAlign) {
        this.vertAlign = vertAlign;
    }

    public void setWesternFontFamily(String westernFontFamily) {
        this.westernFontFamily = westernFontFamily;
    }

    public String getColor() {
        return color;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public double getFontSize() {
        return fontSize;
    }

    public Boolean isBold() {
        return isBold;
    }

    public Boolean isItalic() {
        return isItalic;
    }

    public Boolean isStrike() {
        return isStrike;
    }

    public UnderlinePatterns getUnderlinePatterns() {
        return underlinePatterns;
    }

    public String getUnderlineColor() {
        return underlineColor;
    }

    public XWPFHighlightColor getHighlightColor() {
        return highlightColor;
    }

    public int getCharacterSpacing() {
        return characterSpacing;
    }

    public String getVertAlign() {
        return vertAlign;
    }

    public String getWesternFontFamily() {
        return westernFontFamily;
    }


}
