package org.PierDocx;

import org.PierDocx.style.ParagraphStyle;
import org.PierDocx.style.RunStyle;
import org.PierDocx.utils.StyleUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import static org.PierDocx.utils.PicUtils.get_pic_type;


public class PierParagraph {
    public XWPFParagraph paragraph;
    ArrayList<PierRun> runs = new ArrayList<>();
    int size;

    public ArrayList<PierRun> getRuns() {
        return runs;
    }

    public PierRun getLastRun() {
        return getRuns().get(size - 1);
    }

    public PierParagraph(PierDocument document) {
        super();
        this.paragraph = document.document.createParagraph();
    }

    public PierParagraph(PierTableCell tableCell) {
        super();
        this.paragraph = tableCell.cell.addParagraph();
    }

    public PierRun addRun() {
        PierRun run = new PierRun(this);
        this.runs.add(run);
        this.size += 1;
        return run;
    }

    public PierParagraph addPic(String pic_path, int width, int height, String title, RunStyle title_style) {
        try (InputStream stream = new FileInputStream(pic_path)) {
            this.addStyle(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER));
            this.addRun().addText(title).addStyle(title_style);
            this.addRun().run.addPicture(stream, get_pic_type(pic_path), "Generated", Units.toEMU(width), Units.toEMU(height));
            return this;
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public PierParagraph addStyle(PierParagraph paragraph, ParagraphStyle style) {
        StyleUtils.styleParagraph(paragraph, style);
        return paragraph;
    }

    public PierParagraph addStyle(ParagraphStyle style) {
        StyleUtils.styleParagraph(this, style);
        return this;
    }

    public PierParagraph addStyleById(String style_name) {
        this.addStyle(new ParagraphStyle().setStyleId(style_name));
        return this;
    }

    public PierParagraph addPageBreakBefore(){
        this.addStyle(new ParagraphStyle().setPageBreakBefore(true));
        return this;
    }

    public CTP _getCTP() {
        return this.paragraph.getCTP();
    }


}
