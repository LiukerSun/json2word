package org.PierDocx;

import org.apache.poi.xslf.usermodel.XSLFTextRun;

public class Text {
    private XSLFTextRun text;

    public Text(PierRun run) {
        super();
    }

    public void add_text(String text) {
        this.text.setText(text);
    }
}
