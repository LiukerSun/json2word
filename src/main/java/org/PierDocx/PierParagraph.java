package org.PierDocx;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.util.ArrayList;


public class PierParagraph {
    XWPFParagraph paragraph;
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

    public PierRun add_run() {
        PierRun run = new PierRun(this);
        this.runs.add(run);
        this.size += 1;
        return run;
    }


    public CTP _getCTP() {
        return this.paragraph.getCTP();
    }


}
