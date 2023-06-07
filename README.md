# Json2docx

## Introduction

A way to generate word(docx), based on Apache POI.

## Class &Methods

### Methods

| Method                                  | Parameter | Return   | Description                   |
| --------------------------------------- | --------- | -------- | ----------------------------- |
| `org.PierDocx.utils.JsonUtils.loadJson` | None      | JsonNode | Load JsonNode from Json file. |

### PierDocument

| Class Methods              | Parameter                             | Return                   | Description                       |
| -------------------------- | ------------------------------------- | ------------------------ | --------------------------------- |
| PierDocument (Constructor) | String docxPath (Optional parameters) |                          | Main Document class.              |
| addStyles2temp             | JsonNode                              | None                     | Get last Paragraphs.              |
| addParagraph               | None                                  | PierParagraph            | Add a new Paragraph in Document . |
| addTable                   | int row<br />int col                  | PierTable                |                                   |
| getStyles                  | None                                  | XWPFStyles               | Get all Styles.                   |
| saveDocx                   | String docxPath                       | None                     | Save Document .                   |
| getParagraphs              | None                                  | ArrayList<PierParagraph> | Get all Paragraphs.               |
| getLastParagraph           | None                                  | PierParagraph            | Get last Paragraphs.              |

### PierParagraph

| Class Methods              | Parameter                                                    | Return        | Description                                                  |
| -------------------------- | ------------------------------------------------------------ | ------------- | ------------------------------------------------------------ |
| PierParagraph(Constructor) | PierDocument document<br />OR<br />PierTable.PierTableCell tableCell |               | Can add Paragraph into Document or Table.<br />Table is more important than Paragraph. |
| addPageBreakBefore         | None                                                         | PierParagraph | Add PageBreak Before Paragraph.                              |
| addRun                     | None                                                         | PierRun       | Add a new Run in Paragraph.                                  |
| addPic                     | String picPath<br />int width<br />int height<br />String title<br />RunStyle titleStyle | PierParagraph | Add a new Pic into Paragraph.                                |
| addStyle                   | ParagraphStyle style<br />OR<br />Pier Paragraph paragraph, ParagraphStyle style | PierParagraph | Set Style.                                                   |
| addStyleById               | String styleName                                             | PierParagraph | Set Style by StyleId.                                        |
| getRuns                    | None                                                         |               | Get all Runs in this Paragraph.                              |
| getLastRun                 | None                                                         |               | Get last Run.                                                |

### PierRun

| Class Methods        | Parameter                                                   | Return  | Description                   |
| -------------------- | ----------------------------------------------------------- | ------- | ----------------------------- |
| PierRun(Constructor) | PierParagraph paragraph                                     |         | Can add Run into Paragraph.   |
| addPageBreak         | None                                                        | PierRun | Add a new Page.               |
| addText              | String text                                                 | PierRun | Add Text into Run.            |
| addReturn            | None                                                        | PierRun | Add a Line Break into Run.    |
| addLatex             | String latex                                                | None    | Add a Latex formula into Run. |
| addPic               | String picPath<br />int width<br />int height               | PierRun | Add a new Pic into Run.       |
| addStyle             | PierRun run<br />RunStyle style<br />OR<br />RunStyle style | PierRun | Set Style.                    |

### PierTable

| Class Methods                      | Parameter                                             | Return                  | Description                  |
| ---------------------------------- | ----------------------------------------------------- | ----------------------- | ---------------------------- |
| PierTable(Constructor)             | PierDocument document<br />int rows<br />int cols     |                         | Can add Table into Document. |
| mergeCellsHorizontal               | Integer row<br />Integer fromCell<br />Integer toCell | PierTable               | merge Cells Horizontal.      |
| mergeCellsVertically               | Integer row<br />Integer fromCell<br />Integer toCell | PierTable               |                              |
| getRows                            | None                                                  | ArrayList<PierTableRow> |                              |
| getRow                             | int rowIndex                                          | PierTableRow            |                              |
| PierTableRow.getCells              | None                                                  | ArrayList<PierTableRow> |                              |
| PierTableRow.getCell               | int cellIndex                                         | PierTableRow            |                              |
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
package org.test;

import com.fasterxml.jackson.databind.JsonNode;
import org.PierDocx.PierDocument;
import org.PierDocx.PierTable;
import org.PierDocx.style.ParagraphStyle;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

import static org.PierDocx.utils.JsonUtils.loadJson;


public class Main {
    public static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
        // load config from Json File.
        JsonNode configJson = loadJson("Data/config.json");
        // get tempFilePath and resultFilePath.
        String tempFilePath = configJson.get("templateFile").asText();
        String resultFilePath = configJson.get("resultFile").asText();
        // init Paragraph Styles.
        PierDocument.addStyles2temp(configJson);
        // load Docx from tempFilePath.
        PierDocument document = new PierDocument(tempFilePath);
        document.addParagraph()    // add paragraph
                .addStyleById("title001")      // add style
                .addRun()      // add run
                .addText("Lorem ipsum dolor sit amet")   // add text.
        ;
        document.addParagraph().
                addStyle(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER))
                .addRun()
                .addText("这是一个居中段落。");
        document.addParagraph()
                .addPageBreakBefore()
                .addRun()
                .addLatex("$a^{2}+b^{2}=c^{2}$");
        // add a 3*3 table.
        PierTable table = document.addTable(3, 3);
        // add table style.Table Cell(0,0) add text.
        table.getRow(0).getCell(0).setText("consectetur").setWidth("40%");
        // Merge cell
        table.mergeCellsHorizontal(0, 0, 1);
        table.mergeCellsVertically(0, 1, 2);
        table.getRow(1)
                .getCell(0)
                .setText("这是一个合并单元格")
                .addParagraph()
                .addStyle(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER))
                .addRun()
                .addText("这是表格中的段落");
        table.getRow(1)
                .getCell(1)
                .setWidth("30%")
                .addParagraph()
                .addStyle(new ParagraphStyle().setAlign(ParagraphAlignment.CENTER))
                .addRun()
                .addPic("Data/pics/pic01.jpg", 100, 100);
        // save Docx.
        document.saveDocx(resultFilePath);
    }
}

```

  ![image](https://github.com/LiukerSun/json2word/assets/32071915/88b10716-369d-4539-99f2-2107ab5b1279)


## TODO

- [ ] Table Styles.
- [ ] Load data from json into Table.
- [ ] Feat HyperLink and Bookmark.

