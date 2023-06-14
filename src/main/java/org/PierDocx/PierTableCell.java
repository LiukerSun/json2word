package org.PierDocx;

import org.PierDocx.style.CellStyle;
import org.PierDocx.utils.StyleUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import uk.ac.ed.ph.snuggletex.SnuggleEngine;
import uk.ac.ed.ph.snuggletex.SnuggleInput;
import uk.ac.ed.ph.snuggletex.SnuggleSession;

import java.util.ArrayList;

import static org.PierDocx.utils.LatexUtils._getOMML;

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
    public void addLatex(String latex) throws Exception {
        SnuggleEngine engine = new SnuggleEngine();
        SnuggleSession session = engine.createSession();
        SnuggleInput input = new SnuggleInput(latex);
        session.parseInput(input);
        String mathML = session.buildXMLString();
        CTOMath ctOMath = _getOMML(mathML);
        XWPFParagraph para = this.cell.addParagraph();
        CTP ctp =para.getCTP();
        ctp.setOMathArray(new CTOMath[]{ctOMath});
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
