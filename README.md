# Json2docx

## Introduction

A way to generate word(docx), based on Apache POI.

## Class &Methods

### Methods

| Method                                               | Parameter | Return   | Description                   |
| ---------------------------------------------------- | --------- | -------- | ----------------------------- |
| `org.PierDocx.utils.json.load_json`                  | None      | JsonNode | Load JsonNode from Json file. |
| `org.PierDocx.utils.CreateTempStyles.addStyles2temp` | JsonNode  | None     |                               |

### PierDocument

| Class Methods              | Parameter                              | Return                   | Description                       |
| -------------------------- | -------------------------------------- | ------------------------ | --------------------------------- |
| PierDocument (Constructor) | String docx_path (Optional parameters) |                          | Main Document class.              |
| add_paragraph              | None                                   | PierParagraph            | Add a new Paragraph in Document . |
| add_table                  | int row<br />int col                   | PierTable                |                                   |
| get_styles                 | None                                   | XWPFStyles               | Get all Styles.                   |
| save_docx                  | String docx_path                       | None                     | Save Document .                   |
| getParagraphs              | None                                   | ArrayList<PierParagraph> | Get all Paragraphs.               |
| get_last_paragraph         | None                                   | PierParagraph            | Get last Paragraphs.              |

### PierParagraph

| Class Methods              | Parameter                                                    | Return        | Description                                                  |
| -------------------------- | ------------------------------------------------------------ | ------------- | ------------------------------------------------------------ |
| PierParagraph(Constructor) | PierDocument document<br />OR<br />PierTable.PierTableCell tableCell |               | Can add Paragraph into Document or Table.<br />Table is more important than Paragraph. |
| add_run                    | None                                                         | PierRun       | Add a new Run in Paragraph.                                  |
| getRuns                    | None                                                         |               | Get all Runs in this Paragraph.                              |
| get_last_run               | None                                                         |               | Get last Run.                                                |
| add_pic                    | String pic_path<br />int width<br />int height<br />String title<br />RunStyle title_style | PierParagraph | Add a new Pic into Paragraph.                                |
| add_style                  | ParagraphStyle style<br />OR<br />PierParagraph paragraph, ParagraphStyle style | PierParagraph | Set Style.                                                   |
| add_style_by_id            | String style_name                                            | PierParagraph | Set Style by StyleId.                                        |

### PierRun

| Class Methods        | Parameter                                                   | Return  | Description                   |
| -------------------- | ----------------------------------------------------------- | ------- | ----------------------------- |
| PierRun(Constructor) | PierParagraph paragraph                                     |         | Can add Run into Paragraph.   |
| add_text             | String text                                                 | PierRun | Add Text into Run.            |
| add_return           | None                                                        | PierRun | Add a Line Break into Run.    |
| add_latex            | String latex                                                | None    | Add a Latex formula into Run. |
| add_pic              | String pic_path<br />int width<br />int height              | PierRun | Add a new Pic into Run.       |
| add_style            | PierRun run<br />RunStyle style<br />OR<br />RunStyle style | PierRun | Set Style.                    |

### PierTable

| Class Methods                      | Parameter                                             | Return                  | Description                  |
| ---------------------------------- | ----------------------------------------------------- | ----------------------- | ---------------------------- |
| PierTable(Constructor)             | PierDocument document<br />int rows<br />int cols     |                         | Can add Table into Document. |
| mergeCellsHorizontal               | Integer row<br />Integer fromCell<br />Integer toCell | PierTable               | merge Cells Horizontal.      |
| mergeCellsVertically               | Integer row<br />Integer fromCell<br />Integer toCell | PierTable               |                              |
| get_rows                           | None                                                  | ArrayList<PierTableRow> |                              |
| get_row                            | int row_index                                         | PierTableRow            |                              |
| PierTableRow.get_cells             | None                                                  | ArrayList<PierTableRow> |                              |
| PierTableRow.get_cell              | int cell_index                                        | PierTableRow            |                              |
| PierTableCell.setText              | String text                                           | PierTableCell           |                              |
| PierTableCell.setWidth             | String width                                          | PierTableCell           |                              |
| PierTableCell.setVerticalAlignment | XWPFTableCell.XWPFVertAlign Vert Align                | PierTableCell           |                              |
| PierTableCell.addParagraph         | None                                                  | PierParagraph           |                              |



## Quick start

### Config Example

```json	
{
  "templateFile": "template.docx",
  "resultFile": "result.docx",
  "Styles": [
    {
      "styleName": "leftNormal",
      "font": "宋体",
      "fontSize": 22,
      "alignmentType": "LEFT",
      "isHeading": true,
      "headingLvl": 1,
      "isBold": true,
      "isUnderLine": false,
      "isItalic": false
    },
    {
      "styleName": "rightNormal",
      "font": "宋体",
      "fontSize": 15,
      "alignmentType": "RIGHT",
      "isHeading": false,
      "headingLvl": 1,
      "isBold": false,
      "isUnderLine": false,
      "isItalic": false
    }
  ]
}
```

### Example

```java
import com.fasterxml.jackson.databind.JsonNode;
import org.PierDocx.PierDocument;
import org.PierDocx.PierTable;
import org.PierDocx.style.ParagraphStyle;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import java.io.IOException;
import static org.PierDocx.utils.CreateTempStyles.addStyles2temp;
import static org.PierDocx.utils.json.load_json;

public class Main {
    public static void main(String[] args) throws IOException {
        // load config from Json File.
        JsonNode configJson = load_json("./config.json");
        // get tempFilePath and resultFilePath.
        String tempFilePath = configJson.get("templateFile").asText();
        String resultFilePath = configJson.get("resultFile").asText();
        // init Paragraph Styles.
        addStyles2temp(configJson);

        // load Docx from tempFilePath.
        PierDocument document = new PierDocument(tempFilePath);
        document.add_paragraph()    // add paragraph
                .add_style_by_id("rightNormal")      // add style
                .add_run()      // add run
                .add_text("Lorem ipsum dolor sit amet")   // add text.
        ;
        document.add_paragraph().
                add_style(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER))
                .add_run()
                .add_text("这是一个居中段落。");
        // add a 3*3 table.
        PierTable table = document.add_table(3, 3);
        // Table Cell(0,0) add text.
        table.get_row(0).get_cell(0).setText("consectetur").setWidth("40%");
        // Merge cell
        table.mergeCellsHorizontal(0, 1, 2);
        table.mergeCellsVertically(0, 1, 2);
        table.get_row(1)
                .get_cell(0)
                .setText("这是一个合并单元格")
                .addParagraph()
                .add_run()
                .add_text("这是表格中的段落");
        table.get_row(1)
                .get_cell(1)
                .setWidth("30%")
                .addParagraph()
                .add_style(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER))
                .add_run()
                .add_pic("./test.png", 100, 100);
        // save Docx.
        document.save_docx(resultFilePath);
    }
}
```
  
  ![image](https://github.com/LiukerSun/json2word/assets/32071915/88b10716-369d-4539-99f2-2107ab5b1279)


## TODO

- [ ] Table Styles.
- [ ] Load data from json into Table.
- [ ] Feat HyperLink and Bookmark.

