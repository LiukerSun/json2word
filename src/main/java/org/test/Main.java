package org.test;

import org.PierDocx.PierDocument;
import org.PierDocx.PierParagraph;
import org.PierDocx.style.ParagraphStyle;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

import static org.PierDocx.utils.StyleUtils.*;

public class Main {
    public static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
         String docx_path = "Data/template/template.docx";
//        PierDocument document = new PierDocument();
        PierDocument document = new PierDocument(docx_path);
        PierParagraph paragraph = document.add_paragraph();
        paragraph.add_run().add_latex("$a^{2}+b^{2}=c^{2}$");
        paragraph.add_run().add_text("||||").add_pic("Data/pics/1.jpg",100,100);

        styleParagraph(paragraph,ParagraphStyle.builder().withAlign(ParagraphAlignment.CENTER).build());

        document.save_docx("result.docx");
    }
}
