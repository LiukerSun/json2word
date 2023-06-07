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

    public PierTableCell setText(String text) {
        this.cell.setText(text);
        return this;
    }

    public PierTableCell setWidth(String width) {
        this.cell.setWidth(width);
        return this;
    }

    public void setVerticalAlignment(XWPFTableCell.XWPFVertAlign VertAlign) {
        this.cell.setVerticalAlignment(VertAlign);
    }

    public PierParagraph addParagraph() {
        PierParagraph paragraph = new PierParagraph(this);
        this.paragraphs.add(paragraph);
        this.paragraphs_count += 1;
        return paragraph;
    }


}
