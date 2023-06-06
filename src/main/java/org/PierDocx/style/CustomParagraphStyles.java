package org.PierDocx.style;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;

public class CustomParagraphStyles {
    XWPFDocument document;
    String styleName;
    String fontType;
    String font;
    int fontSize;
    String alignmentType;
    Boolean isHeading;
    int headingLvl;
    Boolean isBold;
    Boolean isUnderLine;
    Boolean isItalic;


    public CustomParagraphStyles(JsonNode styleJson) {
        this.styleName = styleJson.get("styleName").asText();
        this.font = styleJson.get("font").asText();
        this.fontSize = styleJson.get("fontSize").asInt();
        this.alignmentType = styleJson.get("alignmentType").asText();
        this.isHeading = styleJson.get("isHeading").asBoolean();
        if (this.isHeading) {
            this.headingLvl = styleJson.get("headingLvl").asInt();
        }
        this.isBold = styleJson.get("isBold").asBoolean();
        this.isUnderLine = styleJson.get("isUnderLine").asBoolean();
        this.isItalic = styleJson.get("isItalic").asBoolean();
    }


    public XWPFStyle addCustomStyle() {
        // set StyleId.
        CTStyle ctStyle = CTStyle.Factory.newInstance();
        CTString styleName = CTString.Factory.newInstance();
        ctStyle.setStyleId(this.styleName);
        styleName.setVal(this.styleName);
        ctStyle.setName(styleName);
        if (this.isHeading) {
            CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
            indentNumber.setVal(BigInteger.valueOf(this.headingLvl));
            ctStyle.setUiPriority(indentNumber);
            CTPPrGeneral ppr = CTPPrGeneral.Factory.newInstance();
            ppr.setOutlineLvl(indentNumber);
        }

        // set Font FontSize Bold UnderLine
        CTRPr font_ppr = ctStyle.addNewRPr();
        font_ppr.addNewRFonts().setAscii(this.font);
        font_ppr.addNewRFonts().setEastAsia(this.font);
        font_ppr.addNewRFonts().setHAnsi(this.font);
        font_ppr.addNewRFonts().setCs(this.font);
        font_ppr.addNewSz().setVal(this.fontSize * 2);
        font_ppr.addNewSzCs().setVal(this.fontSize * 2);
        ctStyle.setRPr(font_ppr);

        if (this.isBold) {
            font_ppr.addNewB();
        }
        if (this.isUnderLine) {
            font_ppr.addNewU();
            font_ppr.addNewI();
        }
        if (this.isItalic) {
            font_ppr.addNewI();
        }
        // set Paragraph Style
        CTPPrGeneral pPr = ctStyle.addNewPPr();
        CTJc Jc = pPr.addNewJc();
        switch (this.alignmentType.toUpperCase()) {
            case "START" -> Jc.setVal(STJc.Enum.forInt(1));
            case "CENTER" -> Jc.setVal(STJc.Enum.forInt(2));
            case "END" -> Jc.setVal(STJc.Enum.forInt(3));
            case "BOTH" -> Jc.setVal(STJc.Enum.forInt(4));
            case "MEDIUM_KASHIDA" -> Jc.setVal(STJc.Enum.forInt(5));
            case "DISTRIBUTE" -> Jc.setVal(STJc.Enum.forInt(6));
            case "NUM_TAB" -> Jc.setVal(STJc.Enum.forInt(7));
            case "HIGH_KASHIDA" -> Jc.setVal(STJc.Enum.forInt(8));
            case "LOW_KASHIDA" -> Jc.setVal(STJc.Enum.forInt(9));
            case "THAI_DISTRIBUTE" -> Jc.setVal(STJc.Enum.forInt(10));
            case "LEFT" -> Jc.setVal(STJc.Enum.forInt(11));
            case "RIGHT" -> Jc.setVal(STJc.Enum.forInt(12));
        }
        // style shows up in the formats bar
        CTOnOff onOffNull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onOffNull);
        ctStyle.setQFormat(onOffNull);
        // enable Style.
        XWPFStyle style = new XWPFStyle(ctStyle);
        // 这是一个段落样式。
        style.setType(STStyleType.PARAGRAPH);

        return style;

//        XWPFStyles styles = this.document.createStyles();
//        styles.addStyle(style);
    }


}
