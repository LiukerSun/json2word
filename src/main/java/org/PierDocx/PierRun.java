package org.PierDocx;

import org.PierDocx.style.RunStyle;
import org.PierDocx.utils.StyleUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import uk.ac.ed.ph.snuggletex.SnuggleEngine;
import uk.ac.ed.ph.snuggletex.SnuggleInput;
import uk.ac.ed.ph.snuggletex.SnuggleSession;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import static org.PierDocx.utils.LatexUtils._getOMML;
import static org.PierDocx.utils.PicUtils.get_pic_type;

public class PierRun {
    public XWPFRun run;
    PierParagraph paragraph;


    public PierRun(PierParagraph paragraph) {
        this.paragraph = paragraph;
        this.run = paragraph.getParagraph().createRun();
    }

    public PierRun(PierParagraph paragraph, XWPFRun _run) {
        this.paragraph = paragraph;
        this.run = _run;
    }


    public PierRun addText(String text) {
        this.run.setText(text);
        return this;
    }

    public PierRun addReturn() {
        this.run.addCarriageReturn();
        return this;
    }

    public void addLatex(String latex) throws Exception {
        SnuggleEngine engine = new SnuggleEngine();
        SnuggleSession session = engine.createSession();
        SnuggleInput input = new SnuggleInput(latex);
        session.parseInput(input);
        String mathML = session.buildXMLString();
        CTOMath ctOMath = _getOMML(mathML);
        CTP ctp = this.paragraph._getCTP();
        ctp.setOMathArray(new CTOMath[]{ctOMath});
        //        return this;
    }

    public PierRun addPic(String pic_path, int width, int height) {
        try (InputStream stream = new FileInputStream(pic_path)) {
            this.run.addPicture(stream, get_pic_type(pic_path), "Generated", Units.toEMU(width), Units.toEMU(height));
            return this;
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public PierRun addPageBreak() {
        this.run.addBreak(BreakType.PAGE);
        return this;
    }

    public PierRun addStyle(PierRun run, RunStyle style) {
        StyleUtils.styleRun(run, style);
        return this;
    }

    public PierRun addStyle(RunStyle style) {
        StyleUtils.styleRun(this, style);
        return this;
    }

}