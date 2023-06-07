package org.PierDocx.style;

import org.apache.poi.xwpf.usermodel.TableRowAlign;

import java.io.Serializable;

public class TableStyle implements Serializable {
    private TableRowAlign align;

    private BorderStyle leftBorder;
    private BorderStyle rightBorder;
    private BorderStyle topBorder;
    private BorderStyle bottomBorder;
    private BorderStyle insideHBorder;
    private BorderStyle insideVBorder;

    private int leftCellMargin;
    private int topCellMargin;
    private int rightCellMargin;
    private int bottomCellMargin;

    private Double indentation;

    private String width;

    private int[] colWidths;

    // setter
    public TableStyle setAlign(TableRowAlign align) {
        this.align = align;
        return this;
    }

    public void setLeftBorder(BorderStyle leftBorder) {
        this.leftBorder = leftBorder;
    }

    public void setRightBorder(BorderStyle rightBorder) {
        this.rightBorder = rightBorder;
    }

    public void setTopBorder(BorderStyle topBorder) {
        this.topBorder = topBorder;
    }

    public void setBottomBorder(BorderStyle bottomBorder) {
        this.bottomBorder = bottomBorder;
    }

    public void setInsideHBorder(BorderStyle insideHBorder) {
        this.insideHBorder = insideHBorder;
    }

    public void setInsideVBorder(BorderStyle insideVBorder) {
        this.insideVBorder = insideVBorder;
    }

    public TableStyle setWidth(String width) {
        this.width = width;
        return this;
    }

    public TableStyle setColWidths(int[] colWidths) {
        this.colWidths = colWidths;
        return this;
    }

    public TableStyle setLeftCellMargin(int leftCellMargin) {
        this.leftCellMargin = leftCellMargin;
        return this;
    }

    public TableStyle setTopCellMargin(int topCellMargin) {
        this.topCellMargin = topCellMargin;
        return this;
    }

    public TableStyle setRightCellMargin(int rightCellMargin) {
        this.rightCellMargin = rightCellMargin;
        return this;
    }

    public TableStyle setBottomCellMargin(int bottomCellMargin) {
        this.bottomCellMargin = bottomCellMargin;
        return this;
    }

    public TableStyle setIndentation(Double indentation) {
        this.indentation = indentation;
        return this;
    }

    // getter
    public TableRowAlign getAlign() {
        return align;
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

    public BorderStyle getInsideHBorder() {
        return insideHBorder;
    }

    public BorderStyle getInsideVBorder() {
        return insideVBorder;
    }

    public String getWidth() {
        return width;
    }

    public int[] getColWidths() {
        return colWidths;
    }

    public int getLeftCellMargin() {
        return leftCellMargin;
    }

    public int getTopCellMargin() {
        return topCellMargin;
    }

    public int getRightCellMargin() {
        return rightCellMargin;
    }

    public int getBottomCellMargin() {
        return bottomCellMargin;
    }

    public Double getIndentation() {
        return indentation;
    }

}
