package org.PierDocx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;


import java.io.*;
import java.util.ArrayList;

public class PierDocument {
    XWPFDocument document;
    ArrayList<PierParagraph> paragraphs = new ArrayList<>();
    ArrayList<PierTable> tables = new ArrayList<>();
    int paragraphs_count = 0;
    int tables_count = 0;

    public ArrayList<PierParagraph> getParagraphs() {
        return paragraphs;
    }

    public PierParagraph get_last_paragraph() {
        return getParagraphs().get(paragraphs_count - 1);
    }

    public PierDocument(String docx_path) throws IOException {
        InputStream is = new FileInputStream(docx_path);
        this.document = new XWPFDocument(is);
    }

    public PierDocument() {
        this.document = new XWPFDocument();
    }

    public PierParagraph add_paragraph() {
        PierParagraph paragraph = new PierParagraph(this);
        this.paragraphs.add(paragraph);
        this.paragraphs_count += 1;
        return paragraph;
    }

    public PierTable add_table(int row, int col) {
        PierTable table = new PierTable(this, row, col);
        this.tables.add(table);
        this.tables_count += 1;
        return table;
    }

    public void save_docx(String docx_path) throws IOException {
        final FileOutputStream out = new FileOutputStream(docx_path);
        this.document.write(out);
    }

}


