package org.utils;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import static org.main.Main.logger;

public class tools {
    public static JsonNode load_json(String json_path) throws JsonProcessingException {
        StringBuilder sb = new StringBuilder();
        try (InputStream input = new FileInputStream(json_path)) {
            byte[] buffer = new byte[1024];
            int length;
            length = input.read(buffer);
            while (length != -1) {
                sb.append(new String(buffer, 0, length));
                length = input.read(buffer);
            }
        } catch (Exception e) {
            logger.error("load json fail.");
            e.printStackTrace();
        }
        ObjectMapper mapper = new ObjectMapper();
        return mapper.readTree(sb.toString());
    }
    public static void rmWatermark(String docFilePath) {
        try (FileInputStream in = new FileInputStream(docFilePath)) {
            XWPFDocument doc = new XWPFDocument(OPCPackage.open(in));
            List<XWPFParagraph> paragraphs = doc.getParagraphs();
            if (paragraphs.size() < 1) return;
            XWPFParagraph firstParagraph = paragraphs.get(0);
            if (firstParagraph.getText().contains("Spire.Doc")) {
                doc.removeBodyElement(doc.getPosOfParagraph(firstParagraph));
            }
            OutputStream out = new FileOutputStream(docFilePath);
            doc.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
