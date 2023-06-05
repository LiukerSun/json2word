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
import java.io.PipedReader;
import java.util.ArrayList;

import static org.PierDocx.utils.pic.get_pic_type;


public class PierParagraph {
    public XWPFParagraph paragraph;
    ArrayList<PierRun> runs = new ArrayList<>();
    int size;

    public ArrayList<PierRun> getRuns() {
        return runs;
    }

    public PierRun get_last_run() {
        return getRuns().get(size - 1);
    }

    public PierParagraph(PierDocument document) {
        super();
        this.paragraph = document.document.createParagraph();
    }
    public PierParagraph(PierTable.PierTableCell tableCell) {
        super();
        this.paragraph = tableCell.cell.addParagraph();
    }
    public PierRun add_run() {
        PierRun run = new PierRun(this);
        this.runs.add(run);
        this.size += 1;
        return run;
    }

    public PierParagraph add_pic(String pic_path, int width, int height, String title,RunStyle title_style) {
        try (InputStream stream = new FileInputStream(pic_path)) {
            this.add_style(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER));
            this.add_run().add_text(title).add_style(title_style);
            this.add_run().run.addPicture(stream, get_pic_type(pic_path), "Generated", Units.toEMU(width), Units.toEMU(height));
            return this;
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public PierParagraph add_style(PierParagraph paragraph, ParagraphStyle style) {
        StyleUtils.styleParagraph(paragraph, style);
        return paragraph;
    }

    public PierParagraph add_style(ParagraphStyle style) {
        StyleUtils.styleParagraph(this, style);
        return this;
    }

    public CTP _getCTP() {
        return this.paragraph.getCTP();
    }


}
