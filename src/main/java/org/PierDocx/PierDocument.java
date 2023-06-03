package org.PierDocx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;


import java.io.*;
import java.util.ArrayList;

public class PierDocument {
    XWPFDocument document;
    ArrayList<PierParagraph> paragraphs = new ArrayList<>();
    int size = 0;

    public ArrayList<PierParagraph> getParagraphs() {
        return paragraphs;
    }

    public PierParagraph get_last_paragraph() {
        return getParagraphs().get(size - 1);
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
        this.size += 1;
        return paragraph;
    }

    public void save_docx(String docx_path) throws IOException {
        final FileOutputStream out = new FileOutputStream(docx_path);
        this.document.write(out);
    }

}


