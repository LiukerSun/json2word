package org.main;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.spire.doc.Document;
import com.spire.doc.documents.Paragraph;
import org.apache.logging.log4j.LogManager;

import static org.utils.tools.load_json;
import static org.utils.tools.rmWatermark;

public class Main {
    public static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws JsonProcessingException {
        JsonNode config = load_json("./Data/config.json");
        String template_file = config.get("template_file").asText();
        String result_file = config.get("result_file").asText();

        Document doc = new Document();
        doc.loadFromFile(template_file);
        Paragraph p1 =  doc.getLastParagraph();
        p1.setText("test");
        p1.applyStyle("Custom_style01");
        doc.saveToFile(result_file);
        rmWatermark(result_file);
    }
}
