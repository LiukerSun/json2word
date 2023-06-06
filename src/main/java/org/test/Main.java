package org.test;

import com.fasterxml.jackson.databind.JsonNode;
import org.PierDocx.PierDocument;
import org.PierDocx.PierTable;
import org.apache.logging.log4j.LogManager;

import static org.PierDocx.utils.json.load_json;
import static org.PierDocx.utils.CreateTempStyles.addStyles2temp;

public class Main {
    public static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
        JsonNode configJson = load_json("Data/config.json");
        addStyles2temp(configJson);
        String docx_path = configJson.get("templateFile").asText();
//       PierDocument document = new PierDocument();
        PierDocument document = new PierDocument(docx_path);
        document.add_paragraph().add_style_by_id("title001").add_run().add_text("00001");


        document.save_docx("result.docx");
    }
}
