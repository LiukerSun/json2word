package org.test;

import org.PierDocx.PierDocument;
import org.PierDocx.PierParagraph;
import org.apache.logging.log4j.LogManager;

import java.io.IOException;

public class Main {
    public static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
//        String docx_path = "Data/template/template.docx";
        PierDocument document = new PierDocument();

        PierParagraph paragraph = document.add_paragraph();
        paragraph.add_run().add_latex("$a^{2}+b^{2}=c^{2}$");
        paragraph.add_run().add_text("||||").add_pic("Data/pics/1.jpg",100,100);

        document.save_docx("result.docx");
    }
}
