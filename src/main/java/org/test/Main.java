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
        PierDocument.addStyles2temp(configJson,tempFilePath);
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
