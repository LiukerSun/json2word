package org.PierDocx.style;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;

public class CustomParagraphStyles {
    String styleName;
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
        this.headingLvl = styleJson.get("headingLvl").asInt();
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
            ctStyle.setPPr(ppr);
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
            case "START": {
                Jc.setVal(STJc.Enum.forInt(1));
                break;
            }
            case "CENTER": {
                Jc.setVal(STJc.Enum.forInt(2));
                break;
            }
            case "END": {
                Jc.setVal(STJc.Enum.forInt(3));
                break;
            }
            case "BOTH": {
                Jc.setVal(STJc.Enum.forInt(4));
                break;
            }
            case "MEDIUM_KASHIDA": {
                Jc.setVal(STJc.Enum.forInt(5));
                break;
            }
            case "DISTRIBUTE": {
                Jc.setVal(STJc.Enum.forInt(6));
                break;
            }
            case "NUM_TAB": {
                Jc.setVal(STJc.Enum.forInt(7));
                break;
            }
            case "HIGH_KASHIDA": {
                Jc.setVal(STJc.Enum.forInt(8));
                break;
            }
            case "LOW_KASHIDA": {
                Jc.setVal(STJc.Enum.forInt(9));
                break;
            }
            case "THAI_DISTRIBUTE": {
                Jc.setVal(STJc.Enum.forInt(10));
                break;
            }
            case "LEFT": {
                Jc.setVal(STJc.Enum.forInt(11));
                break;
            }
            case "RIGHT": {
                Jc.setVal(STJc.Enum.forInt(12));
                break;
            }
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
    }
}
