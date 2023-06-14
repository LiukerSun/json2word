package org.PierDocx;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import org.PierDocx.style.CustomParagraphStyles;
import org.PierDocx.utils.TOCUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;


import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.test.Main.logger;

public class PierDocument {
    XWPFDocument document;
    ArrayList<PierParagraph> paragraphs = new ArrayList<>();
    ArrayList<PierTable> tables = new ArrayList<>();
    int paragraphs_count = 0;
    int tables_count = 0;
    private CTSdtBlock Sdt;

    public CTSdtBlock getSdt() {
        return Sdt;
    }

    public PierDocument(String docx_path) throws IOException {
        InputStream is = new FileInputStream(docx_path);
        this.document = new XWPFDocument(is);
    }

    public PierDocument() {
        this.document = new XWPFDocument();
    }

    public ArrayList<PierParagraph> getParagraphs() {
        return paragraphs;
    }

    public PierParagraph getLastParagraph() {
        if(this.paragraphs_count==0){
            this.addParagraph();
        }
        return getParagraphs().get(paragraphs_count - 1);
    }

    public static void addStyles2temp(JsonNode configJson) {
        ArrayNode styleArray = (ArrayNode) configJson.get("Styles");
        String templateFile = configJson.get("templateFile").asText();
        try (XWPFDocument document = new XWPFDocument()) {
            XWPFStyles styles = document.createStyles();

            for (JsonNode style : styleArray) {
                styles.addStyle(new CustomParagraphStyles(style).addCustomStyle());
            }
            final FileOutputStream out = new FileOutputStream(templateFile);
            document.write(out);
        } catch (Exception e) {
            logger.error(e);
        }
    }

    public void markTOC() {
        this.Sdt = this.document.getDocument().getBody().addNewSdt();
    }

    public void updateTOC() {
        TOCUtils.addTOC(this);
    }

    public PierParagraph addParagraph() {
        PierParagraph paragraph = new PierParagraph(this);
        this.paragraphs.add(paragraph);
        this.paragraphs_count += 1;
        return paragraph;
    }


    public PierTable addTable(int row, int col) {
        PierTable table = new PierTable(this, row, col);
        this.tables.add(table);
        this.tables_count += 1;
        return table;
    }

    public PierTable addTable(JsonNode tableJson) throws Exception {

        this.addParagraph();
        int tableColSize = tableJson.get("col").asInt();
        int tableRowSize = tableJson.get("row").asInt();
        String title = tableJson.get("title").asText();
        if (!Objects.equals(title, "")) {
            this.addParagraph().addStyleById("tableTitle").addRun().addText(title);
        }

        PierTable table = this.addTable(tableRowSize, tableColSize);
        ArrayNode mergeCellsArray = (ArrayNode) tableJson.get("mergeRules");
        for (JsonNode mergeRule : mergeCellsArray) {
            int firstRow = mergeRule.get("firstRow").asInt();
            int firstColumn = mergeRule.get("firstColumn").asInt();
            int lastRow = mergeRule.get("lastRow").asInt();
            int lastColumn = mergeRule.get("lastColumn").asInt();
            table.mergeCell(firstRow, firstColumn, lastRow, lastColumn);

        }
        ArrayNode tableData = (ArrayNode) tableJson.get("data");
        int tableDataSize = tableData.size();
        for (int i = 0; i < tableDataSize; i++) {
            ArrayNode rowArray = (ArrayNode) tableData.get(i);
            // 更新列表，防止和poi的cells 不匹配。
            table.getRow(i).updateCells();
            for (int colIndex = 0; colIndex < rowArray.size(); colIndex++) {
                // 判定單元格内是否為公式。
                String data = rowArray.get(colIndex).asText();
                PierTableCell tableCell = table.getRow(i).getCell(colIndex);
                tableCell.cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                tableCell.cell.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);

                if (data.contains("$")) {
                    HashMap<String, String> latexMap = new HashMap<>() {
                    };
                    Pattern patternLatex = Pattern.compile(".*?\\$(.*?)\\$.*?");
                    Matcher matcherLatex = patternLatex.matcher(data);
                    while (matcherLatex.find()) {
                        latexMap.put(matcherLatex.group(1), matcherLatex.group(1));
                    }
                    for (String part : data.split("\\$")) {
                        if (latexMap.containsKey(part)) {
                            tableCell.getLastParagraph().addRun().addLatex("$" + part + "$");
                            tableCell.cell.addParagraph().setAlignment(ParagraphAlignment.CENTER);
                        } else {
                            tableCell.getLastParagraph().addRun().addText(part);
                        }
                    }
                } else if (data.contains("%")) {
                    HashMap<String, String> picMap = new HashMap<>() {
                    };
                    Pattern patternPic = Pattern.compile(".*?%(.*?)%.*?");
                    Matcher matcherPic = patternPic.matcher(data);
                    while (matcherPic.find()) {
                        picMap.put(matcherPic.group(1), matcherPic.group(1));
                    }
                    for (String part : data.split("%")) {
                        if (picMap.containsKey(part)) {
                            tableCell.getLastParagraph().addRun().addPic("Data/pics/" + part, 100, 100);
                        } else {
                            tableCell.getLastParagraph().addRun().addText(part);
                        }
                    }
                } else {
                    tableCell.getLastParagraph().addRun().addText(data);
                }

            }
        }
        return table;
    }

    public XWPFStyles getStyles() {
        return this.document.getStyles();
    }

    public void saveDocx(String docx_path) throws IOException {
        final FileOutputStream out = new FileOutputStream(docx_path);
        this.document.write(out);
    }

}


