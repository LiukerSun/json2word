package org.test;

import com.fasterxml.jackson.databind.JsonNode;
import org.PierDocx.PierDocument;

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

        PierDocument document = new PierDocument(tempFilePath);

        document.addTable(tableTest1);


//        document.markTOC();

//        document.addParagraph().addPageBreakBefore().addStyleById("title1").addRun().addText("1  工程概况");
//        document.addParagraph().addStyleById("title2").addRun().addText("1.1  项目概况");
//        document.addParagraph().addRun().addText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Mauris tincidunt fermentum leo, at finibus ligula bibendum vitae. Vestibulum id facilisis urna. Phasellus faucibus leo et urna ullamcorper, id pharetra ligula condimentum. Duis fringilla et ligula vitae viverra. Aenean risus est, rutrum nec imperdiet sed, dictum quis velit. Ut eget ante enim. In id massa vel mauris tristique tempus in id nulla. Mauris finibus magna nisi, quis eleifend arcu eleifend ac. Curabitur ultricies dapibus libero vel bibendum. Duis luctus, velit id porttitor hendrerit, risus eros fermentum ipsum, in tincidunt ex tortor ac nulla. Etiam mollis posuere felis tincidunt pharetra. Pellentesque nec volutpat eros. Etiam luctus elit et leo accumsan, eget lacinia nisi ultrices.");
//
//        document.addParagraph().addStyleById("title2").addRun().addText("1.2  构件编号说明");
//        document.addParagraph().addRun().addText("Mauris iaculis dolor et metus aliquet egestas. Nulla convallis tortor at faucibus lobortis. Vestibulum lacinia egestas dolor. Proin eget lorem placerat, efficitur nunc et, mollis tortor. Sed vel elementum ante. Aliquam mauris est, fermentum vitae tellus vel, elementum tempor orci. Suspendisse vel quam lectus. Praesent tempus facilisis ultrices. Quisque in porttitor massa, eu efficitur mi. Morbi vel dolor at dui interdum mollis. Fusce efficitur orci a elit malesuada, ut fringilla ipsum bibendum. Vestibulum id metus eu libero bibendum fringilla. Ut nisl nisl, egestas vitae vestibulum quis, laoreet non urna.");
//
//        document.addParagraph().addStyleById("title1").addPageBreakBefore().addRun().addText("2 检测内容与方法");
//
//        document.addParagraph().addStyleById("title2").addRun().addText("2.1 码头沉降位移检测");
//        document.addParagraph().addRun().addText("Donec a ipsum in ipsum porta accumsan eu non eros. Quisque suscipit justo arcu, id feugiat orci porttitor et. Sed sit amet eleifend tellus. In non posuere ligula. Aliquam sagittis orci ut fringilla ultrices. Nam varius quam et vestibulum rutrum. Morbi eu luctus risus, non lacinia purus. Aliquam egestas lacus non leo vulputate vulputate. Suspendisse placerat egestas lectus ac lobortis. Praesent pellentesque fermentum dui vitae auctor. Praesent dapibus justo eu ante consectetur vestibulum. Donec eu varius leo.");
//
//        document.addParagraph().addStyleById("title2").addRun().addText("2.2 混凝土结构检测");
//        document.addParagraph().addRun().addText("Mauris sagittis erat sed nibh convallis, sit amet volutpat quam ornare. Suspendisse potenti. Cras pretium bibendum dui eu bibendum. Mauris mattis nibh nisi. Fusce orci odio, interdum at tristique vel, sodales sit amet mauris. Aliquam ac mi vitae metus hendrerit dignissim. Maecenas porta quam blandit ligula imperdiet blandit. Nam placerat convallis augue eu euismod. Nullam pretium sollicitudin quam ut eleifend. Fusce nec euismod massa, ac aliquet metus. Proin orci quam, aliquam sit amet sapien id, pharetra scelerisque turpis. Pellentesque eleifend facilisis augue. Mauris sodales sed elit a pulvinar. Duis tincidunt diam nec velit ultricies, vitae semper lorem aliquam. Donec malesuada dui metus, et dapibus metus tincidunt nec.");
//        document.addParagraph().addRun().addLatex("$a^{2}$");


//        document.addTOC();
        document.saveDocx(resultFilePath);
    }
}
