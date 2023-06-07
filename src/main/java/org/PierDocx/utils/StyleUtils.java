package org.PierDocx.utils;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;

import org.PierDocx.*;
import org.PierDocx.style.*;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;

import org.apache.xmlbeans.SimpleValue;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STHexColorRGB;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

public final class StyleUtils {

    public static void styleParagraph(PierParagraph paragraph, ParagraphStyle style) {
        if (null == paragraph || null == style) return;
        stylePpr(paragraph, style);
        styleParaRpr(paragraph, style.getGlyphStyle());
    }

    private static void styleParaRpr(CTParaRPr pr, Style style) {
        if (null == pr || null == style) return;
        if (StringUtils.isNotBlank(style.getColor())) {
            CTColor color = pr.sizeOfColorArray() > 0 ? pr.getColorArray(0) : pr.addNewColor();
            color.setVal(style.getColor());
        }

        if (null != style.isItalic()) {
            CTOnOff italic = pr.sizeOfIArray() > 0 ? pr.getIArray(0) : pr.addNewI();
            italic.setVal(style.isItalic() ? XWPFOnOff.ON : XWPFOnOff.OFF);
        }

        if (null != style.isBold()) {
            CTOnOff bold = pr.sizeOfBArray() > 0 ? pr.getBArray(0) : pr.addNewB();
            bold.setVal(style.isBold() ? XWPFOnOff.ON : XWPFOnOff.OFF);
        }

        if (0 != style.getFontSize() && -1 != style.getFontSize()) {
            BigDecimal bd = BigDecimal.valueOf(style.getFontSize());
            CTHpsMeasure ctSize = pr.sizeOfSzArray() > 0 ? pr.getSzArray(0) : pr.addNewSz();
            ctSize.setVal(bd.multiply(BigDecimal.valueOf(2)).setScale(0, RoundingMode.HALF_UP).toBigInteger());
        }

        if (null != style.isStrike()) {
            CTOnOff strike = pr.sizeOfStrikeArray() > 0 ? pr.getStrikeArray(0) : pr.addNewStrike();
            strike.setVal(style.isStrike() ? XWPFOnOff.ON : XWPFOnOff.OFF);
        }

        UnderlinePatterns underlinePatern = style.getUnderlinePatterns();
        if (null != underlinePatern) {
            CTUnderline underline = pr.sizeOfUArray() > 0 ? pr.getUArray(0) : pr.addNewU();
            underline.setVal(STUnderline.Enum.forInt(underlinePatern.getValue()));
            if (null != style.getUnderlineColor()) {
                String color = style.getUnderlineColor();
                SimpleValue svColor = null;
                if (color.equals("auto")) {
                    STHexColorAuto hexColor = STHexColorAuto.Factory.newInstance();
                    hexColor.setEnumValue(STHexColorAuto.Enum.forString(color));
                    svColor = (SimpleValue) hexColor;
                } else {
                    STHexColorRGB rgbColor = STHexColorRGB.Factory.newInstance();
                    rgbColor.setStringValue(color);
                    svColor = (SimpleValue) rgbColor;
                }
                underline.setColor(svColor);
            }
        }

        CTFonts fonts = pr.sizeOfRFontsArray() > 0 ? pr.getRFontsArray(0) : pr.addNewRFonts();
        String fontFamily = style.getFontFamily();
        if (StringUtils.isNotBlank(fontFamily)) {
            fonts.setEastAsia(fontFamily);
            fonts.setAscii(fontFamily);
            fonts.setHAnsi(fontFamily);
            fonts.setCs(fontFamily);
        }
        String westernFontFamily = style.getWesternFontFamily();
        if (StringUtils.isNotBlank(westernFontFamily)) {
            fonts.setAscii(westernFontFamily);
            fonts.setHAnsi(westernFontFamily);
            fonts.setCs(westernFontFamily);
        }
    }

    public static void styleParaRpr(PierParagraph paragraph, Style style) {
        if (null == paragraph || null == style) return;
        CTP ctp = paragraph.paragraph.getCTP();
        CTPPr pPr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTParaRPr pr = pPr.isSetRPr() ? pPr.getRPr() : pPr.addNewRPr();
        StyleUtils.styleParaRpr(pr, style);
    }

    public static void stylePpr(PierParagraph paragraph, ParagraphStyle style) {
        if (null == paragraph || null == style) return;
        if (null != style.getAlign()) {
            paragraph.paragraph.setAlignment(style.getAlign());
        }

        if (null != style.getSpacing()) {
            paragraph.paragraph.setSpacingBetween(style.getSpacing(),
                    null == style.getSpacingRule() ? LineSpacingRule.AUTO : style.getSpacingRule());
        }
        if (null != style.getSpacingBeforeLines()) {
            paragraph.paragraph.setSpacingBeforeLines(
                    new BigInteger(String.valueOf(Math.round(style.getSpacingBeforeLines() * 100.0))).intValue());
        }
        if (null != style.getSpacingAfterLines()) {
            paragraph.paragraph.setSpacingAfterLines(
                    new BigInteger(String.valueOf(Math.round(style.getSpacingAfterLines() * 100.0))).intValue());
        }
        if (null != style.getSpacingBefore()) {
            paragraph.paragraph.setSpacingBefore(UnitUtils.point2Twips(style.getSpacingBefore()));
        }
        if (null != style.getSpacingAfter()) {
            paragraph.paragraph.setSpacingAfter(UnitUtils.point2Twips(style.getSpacingAfter()));
        }

        CTP ctp = paragraph.paragraph.getCTP();
        CTPPr pr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTInd indent = pr.isSetInd() ? pr.getInd() : pr.addNewInd();
        if (null != style.getIndentLeftChars()) {
            BigInteger bi = new BigInteger(String.valueOf(Math.round(style.getIndentLeftChars() * 100.0)));
            indent.setLeftChars(bi);
            if (indent.isSetLeft()) indent.unsetLeft();
        }
        if (null != style.getIndentRightChars()) {
            BigInteger bi = new BigInteger(String.valueOf(Math.round(style.getIndentRightChars() * 100.0)));
            indent.setRightChars(bi);
            if (indent.isSetRight()) indent.unsetRight();
        }
        if (null != style.getIndentHangingChars()) {
            BigInteger bi = new BigInteger(String.valueOf(Math.round(style.getIndentHangingChars() * 100.0)));
            indent.setHangingChars(bi);
            if (indent.isSetHanging()) indent.unsetHanging();
        }
        if (null != style.getIndentFirstLineChars()) {
            BigInteger bi = new BigInteger(String.valueOf(Math.round(style.getIndentFirstLineChars() * 100.0)));
            indent.setFirstLineChars(bi);
            if (indent.isSetFirstLine()) indent.unsetFirstLine();
        }

        CTPBdr ct = pr.isSetPBdr() ? pr.getPBdr() : pr.addNewPBdr();
        if (null != style.getLeftBorder()) {
            styleCTBorder(ct.isSetLeft() ? ct.getLeft() : ct.addNewLeft(), style.getLeftBorder());
        }
        if (null != style.getTopBorder()) {
            styleCTBorder(ct.isSetTop() ? ct.getTop() : ct.addNewTop(), style.getTopBorder());
        }
        if (null != style.getRightBorder()) {
            styleCTBorder(ct.isSetRight() ? ct.getRight() : ct.addNewRight(), style.getRightBorder());
        }
        if (null != style.getBottomBorder()) {
            styleCTBorder(ct.isSetBottom() ? ct.getBottom() : ct.addNewBottom(), style.getBottomBorder());
        }

        if (null != style.getBackgroundColor()) {
            CTShd shd = pr.isSetShd() ? pr.getShd() : pr.addNewShd();
            XWPFShadingPattern shadingPattern = style.getShadingPattern();
            if (null == shadingPattern) {
                shd.setVal(STShd.CLEAR);
            } else {
                shd.setVal(STShd.Enum.forInt(shadingPattern.getValue()));
            }
            shd.setColor("auto");
            shd.setFill(style.getBackgroundColor());
        }

        if (null != style.getStyleId()) {
            paragraph.paragraph.setStyle(style.getStyleId());
        }

        if (null != style.getKeepLines()) {
            CTOnOff ctKeepLines = pr.isSetKeepLines() ? pr.getKeepLines() : pr.addNewKeepLines();
            ctKeepLines.setVal(style.getKeepLines() ? XWPFOnOff.ON : XWPFOnOff.OFF);
        }
        if (null != style.getKeepNext()) {
            paragraph.paragraph.setKeepNext(style.getKeepNext());
        }
        if (null != style.getPageBreakBefore()) {
            paragraph.paragraph.setPageBreak(style.getPageBreakBefore());
        }
        if (null != style.getWidowControl()) {
            CTOnOff ctWC = pr.isSetWidowControl() ? pr.getWidowControl() : pr.addNewWidowControl();
            ctWC.setVal(style.getWidowControl() ? XWPFOnOff.ON : XWPFOnOff.OFF);
        }
        if (null != style.getAllowWordBreak()) {
//            paragraph.paragraph.setWordWrapped(style.getWordWrap());
            CTOnOff ctWW = pr.isSetWordWrap() ? pr.getWordWrap() : pr.addNewWordWrap();
            ctWW.setVal(style.getAllowWordBreak() ? XWPFOnOff.OFF : XWPFOnOff.ON);
        }

        if (null != style.getNumId()) {
            paragraph.paragraph.setNumID(BigInteger.valueOf(style.getNumId()));
        }
        if (null != style.getLvl()) {
            paragraph.paragraph.setNumILvl(BigInteger.valueOf(style.getLvl()));
        }
    }

    public static void styleCTBorder(CTBorder b, BorderStyle style) {
        if (null != style.getType()) b.setVal(STBorder.Enum.forString(style.getType().toString().toLowerCase()));
        b.setSz(BigInteger.valueOf(style.getSize()));
        b.setSpace(BigInteger.valueOf(style.getSpace()));
        if (null != style.getColor()) b.setColor(style.getColor());
    }

    public static void styleRun(PierRun run, RunStyle style) {
        if (null == run || null == style) return;
        CTRPr pr = getRunProperties(run.run);
        String color = style.getColor();
        if (StringUtils.isNotBlank(color)) {
            // run.setColor(color);
            // issue 326
            CTColor ctColor = pr.sizeOfColorArray() > 0 ? pr.getColorArray(0) : pr.addNewColor();
            ctColor.setVal(color);
            if (ctColor.isSetThemeColor()) ctColor.unsetThemeColor();
        }
        double fontSize = style.getFontSize();
        if (0 != fontSize && -1 != fontSize) {
            run.run.setFontSize(fontSize);
        }
        String fontFamily = style.getFontFamily();
        if (StringUtils.isNotBlank(fontFamily)) {
            run.run.setFontFamily(fontFamily, XWPFRun.FontCharRange.eastAsia);
            run.run.setFontFamily(fontFamily, XWPFRun.FontCharRange.ascii);
            run.run.setFontFamily(fontFamily, XWPFRun.FontCharRange.hAnsi);
            run.run.setFontFamily(fontFamily, XWPFRun.FontCharRange.cs);
        }
        String westernFontFamily = style.getWesternFontFamily();
        if (StringUtils.isNotBlank(westernFontFamily)) {
            run.run.setFontFamily(westernFontFamily, XWPFRun.FontCharRange.ascii);
            run.run.setFontFamily(westernFontFamily, XWPFRun.FontCharRange.hAnsi);
            run.run.setFontFamily(westernFontFamily, XWPFRun.FontCharRange.cs);
        }
        XWPFHighlightColor highlightColor = style.getHighlightColor();
        if (null != highlightColor) {
            CTHighlight highlight = pr.sizeOfHighlightArray() > 0 ? pr.getHighlightArray(0) : pr.addNewHighlight();
            highlight.setVal(STHighlightColor.Enum.forInt(highlightColor.getValue()));
        }
        Boolean bold = style.isBold();
        if (null != bold) run.run.setBold(bold);
        Boolean italic = style.isItalic();
        if (null != italic) run.run.setItalic(italic);
        Boolean strike = style.isStrike();
        if (null != strike) run.run.setStrikeThrough(strike);
        UnderlinePatterns underlinePatern = style.getUnderlinePatterns();
        if (null != underlinePatern) {
            run.run.setUnderline(underlinePatern);
            if (null != style.getUnderlineColor()) {
                run.run.setUnderlineColor(style.getUnderlineColor());
            }
        }
        int point = style.getCharacterSpacing();
        // in twentieths of a point
        if (0 != point && -1 != point) run.run.setCharacterSpacing(UnitUtils.point2Twips(point));
        String vertAlign = style.getVertAlign();
        if (StringUtils.isNotBlank(vertAlign)) {
            run.run.setVerticalAlignment(vertAlign);
        }
    }

    public static void styleTable(PierTable table, TableStyle tableStyle) {
        BorderStyle bottomBorder = tableStyle.getBottomBorder();
        BorderStyle leftBorder = tableStyle.getLeftBorder();
        BorderStyle rightBorder = tableStyle.getRightBorder();
        BorderStyle topBorder = tableStyle.getTopBorder();
        BorderStyle insideHBorder = tableStyle.getInsideHBorder();
        BorderStyle insideVBorder = tableStyle.getInsideVBorder();
        if (null == table || null == tableStyle) return;
        if (null != bottomBorder) {
            table.table.setBottomBorder(
                    bottomBorder.getType(),
                    bottomBorder.getSize(),
                    bottomBorder.getSpace(),
                    bottomBorder.getColor()
            );
        }
        if (null != leftBorder) {
            table.table.setLeftBorder(
                    leftBorder.getType(),
                    leftBorder.getSize(),
                    leftBorder.getSpace(),
                    leftBorder.getColor()
            );
        }
        if (null != rightBorder) {
            table.table.setRightBorder(
                    rightBorder.getType(),
                    rightBorder.getSize(),
                    rightBorder.getSpace(),
                    rightBorder.getColor()
            );
        }
        if (null != topBorder) {
            table.table.setTopBorder(
                    topBorder.getType(),
                    topBorder.getSize(),
                    topBorder.getSpace(),
                    topBorder.getColor()
            );
        }
        if (null != insideHBorder) {
            table.table.setInsideHBorder(
                    insideHBorder.getType(),
                    insideHBorder.getSize(),
                    insideHBorder.getSpace(),
                    insideHBorder.getColor()
            );
        }
        if (null != insideVBorder) {
            table.table.setInsideVBorder(
                    insideVBorder.getType(),
                    insideVBorder.getSize(),
                    insideVBorder.getSpace(),
                    insideVBorder.getColor()
            );
        }

        if (null != tableStyle.getAlign()) {
            table.table.setTableAlignment(tableStyle.getAlign());
        }
        table.table.setCellMargins(tableStyle.getTopCellMargin(), tableStyle.getLeftCellMargin(),
                tableStyle.getBottomCellMargin(), tableStyle.getRightCellMargin());
        if (null != tableStyle.getIndentation()) {
            CTTbl ctTbl = table.table.getCTTbl();
            CTTblPr tPr = (ctTbl.getTblPr() != null) ? ctTbl.getTblPr() : ctTbl.addNewTblPr();
            CTTblWidth tw = tPr.isSetTblInd() ? tPr.getTblInd() : tPr.addNewTblInd();
            tw.setType(STTblWidth.DXA);
            tw.setW(BigInteger.valueOf(UnitUtils.cm2Twips(tableStyle.getIndentation())));
        }
    }

    public static void styleTableRow(PierTableRow row, RowStyle rowStyle) {
        if (null == row || null == rowStyle) return;
        int height = rowStyle.getHeight();
        CTRow ctRow = row.row.getCtRow();
        CTTrPr properties = (ctRow.isSetTrPr()) ? ctRow.getTrPr() : ctRow.addNewTrPr();
        if (0 != height) {
            row.row.setHeight(height);
            CTHeight h = properties.sizeOfTrHeightArray() == 0 ? properties.addNewTrHeight()
                    : properties.getTrHeightArray(0);
            String heightRule = rowStyle.getHeightRule();
            if ("exact".equals(heightRule))
                h.setHRule(STHeightRule.EXACT);
            else if ("atleast".equals(heightRule))
                h.setHRule(STHeightRule.AT_LEAST);
            else
                h.setHRule(STHeightRule.AUTO);
        }

        boolean repeated = rowStyle.isRepeated();
        if (repeated) {
            CTOnOff tblHeader = properties.sizeOfTblHeaderArray() == 0 ? properties.addNewTblHeader()
                    : properties.getTblHeaderArray(0);
            tblHeader.setVal(XWPFOnOff.ON);
        }

        boolean breakAcrossPage = rowStyle.isBreakAcrossPage();
        if (!breakAcrossPage) {
            if (properties.sizeOfCantSplitArray() == 0) {
                properties.addNewCantSplit();
            } else {
                properties.getCantSplitArray(0).setVal(XWPFOnOff.ON);
            }
        }
    }

    public static void styleTableCell(PierTableCell cell, CellStyle cellStyle) {
        if (null == cell || null == cellStyle) return;
        if (null != cellStyle.getVertAlign()) {
            cell.setVerticalAlignment(cellStyle.getVertAlign());
        }
        if (null != cellStyle.getBackgroundColor()) {
            CTTc ctTc = cell.cell.getCTTc();
            CTTcPr pr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
            CTShd shd = pr.isSetShd() ? pr.getShd() : pr.addNewShd();
            XWPFShadingPattern shadingPattern = cellStyle.getShadingPattern();
            if (null == shadingPattern) {
                shd.setVal(STShd.CLEAR);
            } else {
                shd.setVal(STShd.Enum.forInt(shadingPattern.getValue()));
            }
            shd.setColor("auto");
            shd.setFill(cellStyle.getBackgroundColor());
        }
    }



    private static CTRPr getRunProperties(XWPFRun run) {
        return run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
    }
}


