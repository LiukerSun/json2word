package org.test;

import org.PierDocx.PierDocument;
import org.PierDocx.PierTable;
import org.PierDocx.style.ParagraphStyle;
import org.PierDocx.style.RunStyle;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class Main {
    public static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
        String docx_path = "Data/template/template.docx";
        // PierDocument document = new PierDocument();
        PierDocument document = new PierDocument(docx_path);
        document.add_paragraph().add_style(new ParagraphStyle().setAlign(ParagraphAlignment.RIGHT))
                .add_run()
                .add_text("BSYP07001H")
                .add_return()
                .add_text("SY（2022）第**号")
                .add_return()
                .add_return()
                .add_return()
                .add_style(new RunStyle()
                        .setFontSize(12)
                        .setFontFamily("Times New Roman"));

        document.add_paragraph().add_style(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER))
                .add_run()
                .add_text("检 测 报 告")
                .add_page_break()
                .add_style(new RunStyle()
                        .setBold(true)
                        .setFontSize(22)
                        .setFontFamily("宋体"))
        ;

//        document.add_paragraph().add_pic("Data/pics/pic01.jpg", 410, 650, "单位经营范围及资质", new RunStyle().setBold(true).setFontSize(16));

        PierTable table = document.add_table(4, 4).mergeCellsHorizontal(0, 0, 1).mergeCellsVertically(0, 1, 2);
        table.get_row(0).get_cell(0).setText("cell直接添加text").addParagraph().add_run().add_text("cell中插入paragraph");
        document.save_docx("result.docx");
    }
}
