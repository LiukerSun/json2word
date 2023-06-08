package org.PierDocx;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import org.PierDocx.style.CustomParagraphStyles;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFStyles;


import java.io.*;
import java.util.ArrayList;

import static org.test.Main.logger;

public class PierDocument {
    XWPFDocument document;
    ArrayList<PierParagraph> paragraphs = new ArrayList<>();
    ArrayList<PierTable> tables = new ArrayList<>();
    int paragraphs_count = 0;
    int tables_count = 0;

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

    public XWPFStyles getStyles() {
        return this.document.getStyles();
    }

    public void saveDocx(String docx_path) throws IOException {
        final FileOutputStream out = new FileOutputStream(docx_path);
        this.document.write(out);
    }

}


