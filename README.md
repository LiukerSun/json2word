# Json2docx

## Introduction

A way to generate word(docx), based on Apache POI.

## Class &Methods

### Methods

| Method                                  | Parameter | Return   | Description                   |
| --------------------------------------- | --------- | -------- | ----------------------------- |
| `org.PierDocx.utils.JsonUtils.loadJson` | None      | JsonNode | Load JsonNode from Json file. |

### PierDocument

| Class Methods              | Parameter                             | Return                   | Description                                          |
| -------------------------- | ------------------------------------- | ------------------------ | ---------------------------------------------------- |
| PierDocument (Constructor) | String docxPath (Optional parameters) |                          | Main Document class.                                 |
| addStyles2temp             | JsonNode                              | None                     | Get last Paragraphs.                                 |
| addParagraph               | None                                  | PierParagraph            | Add a new Paragraph in Document .                    |
| addTable                   | int row<br />int col                  | PierTable                |                                                      |
| getStyles                  | None                                  | XWPFStyles               | Get all Styles.                                      |
| saveDocx                   | String docxPath                       | None                     | Save Document .                                      |
| getParagraphs              | None                                  | ArrayList<PierParagraph> | Get all Paragraphs.                                  |
| getLastParagraph           | None                                  | PierParagraph            | Get last Paragraphs.                                 |
| markTOC                    | None                                  | None                     | Mark TOC position.                                   |
| ~~updateTOC~~(Deprecated)  | ~~None~~                              | ~~None~~                 | ~~Update TOC.<br />Function Error.Now Cant update.~~ |

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

Add `tableTitle`  for table caption.

```json
{
  "templateFile": "template.docx",
  "resultFile": "result.docx",
  "Styles": [
    {
      "styleName": "tableTitle",
      "font": "黑体",
      "fontSize": 10,
      "alignmentType": "CENTER",
      "isHeading": false,
      "headingLvl": 2,
      "isBold": false,
      "isUnderLine": false,
      "isItalic": false
    },
    {
      "styleName": "title1",
      "font": "宋体",
      "fontSize": 18,
      "alignmentType": "LEFT",
      "isHeading": true,
      "headingLvl": 1,
      "isBold": true,
      "isUnderLine": false,
      "isItalic": false
    }
  ],
  "tables": {
    "test1": {
      "title": "table_test",
      "col": 12,
      "row": 18,
      "mergeRules": [
        {
          "firstRow": 0,
          "firstColumn": 0,
          "lastRow": 1,
          "lastColumn": 0
        },
        {
          "firstRow": 0,
          "firstColumn": 1,
          "lastRow": 1,
          "lastColumn": 1
        },
        {
          "firstRow": 0,
          "firstColumn": 2,
          "lastRow": 0,
          "lastColumn": 11
        },
        {
          "firstRow": 2,
          "firstColumn": 0,
          "lastRow": 16,
          "lastColumn": 0
        },
        {
          "firstRow": 17,
          "firstColumn": 2,
          "lastRow": 17,
          "lastColumn": 11
        },
        {
          "firstRow": 17,
          "firstColumn": 0,
          "lastRow": 17,
          "lastColumn": 1
        }
      ],
      "data": [
        ["构件名称","构件编号","各测区混凝土强度换算值（MPa）"],
        ["","","1","2","3","4","5","6","7","8","9","10"],
        ["前边梁","1#QBL3-4","47.6","46.7","43.4","43.1","41.3","40.4","38.6","36.1","38.1","37"],
        ["","1#QBL5-6","35.8","39.3","37.4","37.9","40.1","41.6","43.4","38.6","39.9","38.2"],
        ["","1#QBL7-8","43.4","42.8","41.1","43.4","40","40.8","40.6","40","38.6","39.5"],
        ["","1#QBL9-10","36","36.8","37.9","36.3","38.8","37.5","39.2","37.4","39.9","43.1"],
        ["","1#QBL11-12","39.9","41.3","39","40.2","40.8","36.1","37.5","38.2","39.9","41.9"],
        ["","2#QBL3-4","42.5","42.9","43.2","43.2","40.2","41","42.1","40.4","40.6","42.5"],
        ["","2#QBL5-6","42","43.9","45.7","45.3","45.3","43.2","43.7","42","43.7","47.3"],
        ["","2#QBL7-8","41.9","42.7","43.6","43.2","40.4","38.5","39","40.4","40.2","43.1"],
        ["","2#QBL9-10","38.8","38.8","37.7","39.7","39.5","40.5","41.2","44.9","44.7","43.5"],
        ["","2#QBL11-12","41.7","39.5","39.9","39","38.6","37.7","39.5","41.7","40.8","41.7"],
        ["","3#QBL3-4","41.9","43.9","38.4","37.5","38.3","43.5","41","43.9","43.4","44.1"],
        ["","3#QBL5-6","43.9","45.7","41.6","44.6","44.7","45.9","44.1","45.1","40.5","44.7"],
        ["","3#QBL7-8","39.3","42.8","43.1","43.7","42.4","46.2","44.3","42.8","41","41.9"],
        ["","3#QBL9-10","44.4","40.4","42.7","38.5","41.2","43.6","41.6","45.8","40.8","42.5"],
        ["","3#QBL11-12","44.8","45","41.6","47","47.1","44","43.6","44.2","45.8","42.9"],
        ["该批构件混凝土强度推定值","批构件强度换算值平均值$a^{2}$,批构件强度换算值平均值$b^{2}$"]
      ]
    },
    "test2": {
      "title": "",
      "col": 3,
      "row": 8,
      "mergeRules": [
        {
          "firstRow": 0,
          "firstColumn": 0,
          "lastRow": 0,
          "lastColumn": 2
        },
        {
          "firstColumn": 0,
          "lastColumn": 0,
          "firstRow": 2,
          "lastRow": 7
        }
      ],
      "data": [
        ["title"],
        ["index","name","number"],
        ["1","liukersun","1"],
        ["","2","2"],
        ["","3","3"],
        ["","%pic01.jpg%","4"],
        ["","5","5"],
        ["","6","6"]
      ]
    }
  }
}
```

### Example

```java
package org.test;

import com.fasterxml.jackson.databind.JsonNode;
import org.PierDocx.PierDocument;

import org.PierDocx.PierTable;
import org.apache.logging.log4j.LogManager;

import static org.PierDocx.utils.JsonUtils.loadJson;


public class Main {
    public static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
        JsonNode configJson = loadJson("Data/config.json");
        String tempFilePath = configJson.get("templateFile").asText();
        String resultFilePath = configJson.get("resultFile").asText();
        PierDocument.addStyles2temp(configJson);
        JsonNode tableTest1 = configJson.get("tables").get("test1");
        JsonNode tableTest2 = configJson.get("tables").get("test2");
        PierDocument document = new PierDocument(tempFilePath);
        //add Paragraph and Title
        document.addParagraph().addPageBreakBefore().addStyleById("title1").addRun().addText("1  Function");
        document.addParagraph().addStyleById("title2").addRun().addText("1.1  Table");
        document.addParagraph().addRun().addText("add table with Latex and table caption.");
        // table with Latex
        document.addTable(tableTest1);
        document.addParagraph().addRun().addText("table with pic and without table caption.");
        // table with pic
        document.addTable(tableTest2);

        document.addParagraph().addPageBreakBefore().addStyleById("title2").addRun().addText("1.2  Latex");
        document.addParagraph().addRun().addText("You can add Latex into word like : ").addReturn().addText("document.addParagraph().addRun().addLatex(\"$a^{2}+b^{2}=c^{2}$\");");
        document.addParagraph().addRun().addLatex("$a^{2}+b^{2}=c^{2}$");
        document.addParagraph().addStyleById("title1").addPageBreakBefore().addRun().addText("2 Quisque");
        document.addParagraph().addStyleById("title2").addRun().addText("2.1 porttitor");
        document.addParagraph().addRun().addText("Donec a ipsum in ipsum porta accumsan eu non eros. Quisque suscipit justo arcu, id feugiat orci porttitor et. Sed sit amet eleifend tellus. In non posuere ligula. Aliquam sagittis orci ut fringilla ultrices. Nam varius quam et vestibulum rutrum. Morbi eu luctus risus, non lacinia purus. Aliquam egestas lacus non leo vulputate vulputate. Suspendisse placerat egestas lectus ac lobortis. Praesent pellentesque fermentum dui vitae auctor. Praesent dapibus justo eu ante consectetur vestibulum. Donec eu varius leo.");
        document.addParagraph().addStyleById("title2").addRun().addText("2.2 sagittis");
        document.addParagraph().addRun().addText("Mauris sagittis erat sed nibh convallis, sit amet volutpat quam ornare. Suspendisse potenti. Cras pretium bibendum dui eu bibendum. Mauris mattis nibh nisi. Fusce orci odio, interdum at tristique vel, sodales sit amet mauris. Aliquam ac mi vitae metus hendrerit dignissim. Maecenas porta quam blandit ligula imperdiet blandit. Nam placerat convallis augue eu euismod. Nullam pretium sollicitudin quam ut eleifend. Fusce nec euismod massa, ac aliquet metus. Proin orci quam, aliquam sit amet sapien id, pharetra scelerisque turpis. Pellentesque eleifend facilisis augue. Mauris sodales sed elit a pulvinar. Duis tincidunt diam nec velit ultricies, vitae semper lorem aliquam. Donec malesuada dui metus, et dapibus metus tincidunt nec.");
        //add pic
        document.addParagraph().addPageBreakBefore().addStyleById("title1").addRun().addText("3 Picture");
        document.addParagraph().addPic("Data/pics/pic01.jpg",100,100,"My WeChat Profile picture","tableTitle");
        document.saveDocx(resultFilePath);
    }
}

```


## TODO

- [ ] add Hyperlink or Bookmark.
- [ ] add TOC. (update next version)
- [ ] add Footer or Header . (update next version)

