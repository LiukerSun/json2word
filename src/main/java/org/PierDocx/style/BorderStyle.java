package org.PierDocx.style;

import java.io.Serializable;

import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;

public class BorderStyle implements Serializable {
    private int size;
    private String color;
    private XWPFBorderType type;
    private int space = 0;

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public int getSize() {
        return size;
    }

    public void setSize(int size) {
        this.size = size;
    }

    public XWPFBorderType getType() {
        return type;
    }

    public void setType(XWPFBorderType type) {
        this.type = type;
    }

    public int getSpace() {
        return space;
    }

    public void setSpace(int space) {
        this.space = space;
    }

    public BorderStyle setDefaultBorderStyle() {
        this.setSize(8 * 1 / 2);
        this.setColor("auto");
        this.setType(XWPFBorderType.SINGLE);
        return this;
    }

}
