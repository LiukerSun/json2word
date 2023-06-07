package org.test;

import com.fasterxml.jackson.databind.JsonNode;
import org.PierDocx.PierDocument;
import org.PierDocx.PierTable;
import org.PierDocx.style.ParagraphStyle;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

import static org.PierDocx.utils.JsonUtils.loadJson;


public class Main {
    public static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
        logger.info("Start");
        // load config from Json File.
        JsonNode configJson = loadJson("Data/config.json");
        // get tempFilePath and resultFilePath.
        String tempFilePath = configJson.get("templateFile").asText();
        String resultFilePath = configJson.get("resultFile").asText();
        // init Paragraph Styles.
        PierDocument.addStyles2temp(configJson);
        // load Docx from tempFilePath.
        PierDocument document = new PierDocument(tempFilePath);
        document.addParagraph()    // add paragraph
                .addStyleById("title001")      // add style
                .addRun()      // add run
                .addText("Lorem ipsum dolor sit amet")   // add text.
        ;
        document.addParagraph().
                addStyle(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER))
                .addRun()
                .addText("这是一个居中段落。");
        document.addParagraph()
                .addPageBreakBefore()
                .addRun()
                .addLatex("$a^{2}+b^{2}=c^{2}$");
        // add a 3*3 table.
        PierTable table = document.addTable(3, 3);
        // add table style.Table Cell(0,0) add text.
        table.getRow(0).getCell(0).setText("consectetur").setWidth("40%");
        // Merge cell
        table.mergeCellsHorizontal(0, 0, 1);
        table.mergeCellsVertically(0, 1, 2);
        table.getRow(1)
                .getCell(0)
                .setText("这是一个合并单元格")
                .addParagraph()
                .addStyle(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER))
                .addRun()
                .addText("这是表格中的段落");
        table.getRow(1)
                .getCell(1)
                .setWidth("30%")
                .addParagraph()
                .addStyle(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER))
                .addRun()
                .addPic("Data/pics/pic01.jpg", 100, 100);
        // save Docx.
        document.saveDocx(resultFilePath);
    }
}
