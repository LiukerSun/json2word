package org.PierDocx.style;

import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;

import java.io.Serializable;

public class CellStyle implements Serializable {
    private String backgroundColor;
    private XWPFShadingPattern shadingPattern;
    private XWPFVertAlign vertAlign;

    // setter
    public void setBackgroundColor(String backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    public void setShadingPattern(XWPFShadingPattern shadingPattern) {
        this.shadingPattern = shadingPattern;
    }

    public void setVertAlign(XWPFVertAlign align) {
        this.vertAlign = align;
    }

    // getter
    public String getBackgroundColor() {
        return backgroundColor;
    }

    public XWPFShadingPattern getShadingPattern() {
        return shadingPattern;
    }

    public XWPFVertAlign getVertAlign() {
        return vertAlign;
    }


}
