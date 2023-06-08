package org.test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import org.PierDocx.PierDocument;
import org.PierDocx.PierTable;
import org.PierDocx.PierTableCell;
import org.PierDocx.style.ParagraphStyle;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

import static org.PierDocx.utils.JsonUtils.loadJson;


public class Main {
    public static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
        JsonNode configJson = loadJson("Data/config.json");
        String tempFilePath = configJson.get("templateFile").asText();
        String resultFilePath = configJson.get("resultFile").asText();
        PierDocument.addStyles2temp(configJson);
        PierDocument document = new PierDocument(tempFilePath);


        ArrayNode tableArray = (ArrayNode) configJson.get("table");
        for (JsonNode tableElement : tableArray) {
            document.addTable(tableElement);
        }

        // save Docx.
        document.saveDocx(resultFilePath);
    }
}
