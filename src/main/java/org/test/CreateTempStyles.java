package org.test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;

import org.PierDocx.style.CustomParagraphStyles;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyles;

import java.io.FileOutputStream;

import static org.test.Main.logger;

public class CreateTempStyles {
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
}
