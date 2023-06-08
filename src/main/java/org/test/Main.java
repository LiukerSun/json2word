package org.test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import org.PierDocx.PierDocument;
import org.PierDocx.PierRun;
import org.PierDocx.PierTable;
import org.PierDocx.PierTableCell;
import org.PierDocx.style.CellStyle;
import org.PierDocx.style.ParagraphStyle;
import org.PierDocx.style.RunStyle;
import org.PierDocx.style.TableStyle;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.Cell;
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
            int tableColSize = tableElement.get("col").asInt();
            int tableRowSize = tableElement.get("row").asInt();
            PierTable table = document.addTable(tableRowSize, tableColSize);
            ArrayNode mergeCellsArray = (ArrayNode) tableElement.get("mergeCells");
            for (JsonNode mergeRule : mergeCellsArray) {
                int firstRow = mergeRule.get("firstRow").asInt();
                int firstColumn = mergeRule.get("firstColumn").asInt();
                int lastRow = mergeRule.get("lastRow").asInt();
                int lastColumn = mergeRule.get("lastColumn").asInt();
                table.mergeCell(firstRow, firstColumn, lastRow, lastColumn);
            }

            ArrayNode tableData = (ArrayNode) tableElement.get("data");
            int tableDataSize = tableData.size();
            for (int i = 0; i < tableDataSize; i++) {
                ArrayNode rowArray = (ArrayNode) tableData.get(i);
                for (int colIndex = 0; colIndex < rowArray.size(); colIndex++) {
                    PierTableCell tableCell = table.getRow(i).getCell(colIndex);
                    tableCell.setText(rowArray.get(colIndex).asText()).addStyle(
                            new ParagraphStyle().setAlign(ParagraphAlignment.CENTER));
                }
            }
        }

        // save Docx.
        document.saveDocx(resultFilePath);
    }
}
