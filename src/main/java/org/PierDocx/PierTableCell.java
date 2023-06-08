package org.PierDocx;

import org.PierDocx.style.CellStyle;
import org.PierDocx.utils.StyleUtils;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import java.util.ArrayList;

public class PierTableCell {
    public XWPFTableCell cell;
    ArrayList<PierParagraph> paragraphs = new ArrayList<>();
    int paragraphs_count = 0;


    public PierTableCell(XWPFTableCell cell) {
        this.cell = cell;
    }

    public PierTableCell addStyle(PierTableCell cell, CellStyle style) {
        StyleUtils.styleTableCell(cell, style);
        return cell;
    }

    public PierTableCell addStyle(CellStyle style) {
        StyleUtils.styleTableCell(this, style);
        return this;
    }

    public PierParagraph setText(String text) {
        // 每次执行setText 都会重置表格的段落数。
        this.cell.setText(text);

        this.paragraphs = new ArrayList<>();
        PierParagraph paragraph = new PierParagraph(this);
        this.paragraphs.add(paragraph);
        this.paragraphs_count = 1;
        return paragraph;
    }

    public PierTableCell setWidth(String width) {
        this.cell.setWidth(width);
        return this;
    }

    public void setVerticalAlignment(XWPFTableCell.XWPFVertAlign VertAlign) {
        this.cell.setVerticalAlignment(VertAlign);
    }

    public PierParagraph addParagraph() {
        PierParagraph paragraph = new PierParagraph(this.cell.getCTTc().addNewP(), this.cell);
        this.paragraphs.add(paragraph);
        this.paragraphs_count += 1;
        return paragraph;
    }

    public PierParagraph getLastParagraph() {
        if (paragraphs_count == 0) {
            return new PierParagraph(this);
        } else {
            return this.paragraphs.get(paragraphs_count - 1);
        }
    }

    public PierParagraph getParagraph(int paragraphsIndex) {
        return this.paragraphs.get(paragraphsIndex);
    }

}
