package org.PierDocx.utils;

import org.PierDocx.PierDocument;
import org.PierDocx.PierParagraph;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtDocPart;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabTlc;

import java.math.BigInteger;
import java.util.ArrayList;


public class TOCUtils {

//    public static void addTOC(PierDocument document) {
//        CTSdtBlock Sdt = document.getSdt();
//        addTitle(Sdt);
//        List<PierParagraph> paragraphs = document.getParagraphs();
//        for (PierParagraph par : paragraphs) {
//            String parStyle = par.getStyleID();
//            if (parStyle != null && parStyle.startsWith("title")) {
////                List<CTBookmark> bookmarkList = par._getCTP().getBookmarkStartList();
//                try {
//                    int level = Integer.parseInt(parStyle.substring("title".length()));
//                    if (level == 1) {
//                        //添加栏目
//                        addRowOnlyTitle(Sdt, level, par.getParagraph().getText());
//                    } else if(level == 2){
//                        //添加标题
//                        addRow(Sdt, level, par.getParagraph().getText(), 1, "");
//                    }else{
//                        addRow(Sdt, level, par.getParagraph().getText(), 1, "");
//                    }
//                } catch (NumberFormatException e) {
//                    e.printStackTrace();
//                }
//            }
//        }
//    }
//
//    private static void addTitle(CTSdtBlock Sdt) {
//        CTSdtPr sdtPr = Sdt.addNewSdtPr();
//        sdtPr.addNewDocPartObj().addNewDocPartGallery().setVal("Table of contents");
//        CTSdtEndPr sdtEndPr = Sdt.addNewSdtEndPr();
//        CTRPr rPr = sdtEndPr.addNewRPr();
//        CTFonts fonts = rPr.addNewRFonts();
//        fonts.setAsciiTheme(STTheme.MINOR_H_ANSI);
//        fonts.setEastAsiaTheme(STTheme.MINOR_H_ANSI);
//        fonts.setHAnsiTheme(STTheme.MINOR_H_ANSI);
//        fonts.setCstheme(STTheme.MINOR_BIDI);
//        rPr.addNewB();
//        rPr.addNewBCs();
//        rPr.addNewColor().setVal("auto");
//        rPr.addNewSz().setVal("24");
//        rPr.addNewSzCs().setVal("24");
//        CTSdtContentBlock content = Sdt.addNewSdtContent();
//        CTP p = content.addNewP();
//        p.addNewPPr().addNewPStyle().setVal("TOCHeading");
//        p.addNewR().addNewT().setStringValue("目     录");//源码中为"Table of contents"
//        //设置段落对齐方式，即将“目录”二字居中
//        CTPPr pr = p.getPPr();
//        CTJc jc = pr.isSetJc() ? pr.getJc() : pr.addNewJc();
//        STJc.Enum en = STJc.Enum.forInt(ParagraphAlignment.CENTER.getValue());
//        jc.setVal(en);
//        //"目录"二字的字体
//        CTRPr pRpr = p.getRArray(0).addNewRPr();
//        fonts = pRpr.addNewRFonts();
//        fonts.setAscii("Times New Roman");
//        fonts.setEastAsia("宋体");
//        fonts.setHAnsi("宋体");
//        //"目录"二字加粗
//        CTOnOff bold = pRpr.addNewB();
//        // 设置“目录”二字字体大小为24号
//        CTHpsMeasure sz = pRpr.addNewSz();
//        sz.setVal("32");
//    }
//
//    public static void addRowOnlyTitle(CTSdtBlock Sdt, int level, String title) {
//        CTSdtContentBlock contentBlock = Sdt.getSdtContent();
//        CTP p = contentBlock.addNewP();
//        CTPPr pPr = p.addNewPPr();
//        pPr.addNewPStyle().setVal("TOC" + level);
//        CTTabs tabs = pPr.addNewTabs();//Set of Custom Tab Stops自定义制表符集合
//        CTTabStop tab = tabs.addNewTab();//Custom Tab Stop自定义制表符
//        tab.setVal(STTabJc.RIGHT);
//        tab.setLeader(STTabTlc.DOT);
//        tab.setPos("8200");//默认为8290，因为调整过页边距，所有需要调整，手动设置找出最佳值
//        pPr.addNewRPr().addNewNoProof();//不检查语法
//        CTR run = p.addNewR();
//        run.addNewRPr().addNewNoProof();
//        run.addNewT().setStringValue(title);
//        //设置行间距
//        CTSpacing pSpacing = pPr.getSpacing() != null ? pPr.getSpacing() : pPr.addNewSpacing();
//        pSpacing.setLineRule(STLineSpacingRule.AUTO);//行间距类型：多倍
//        pSpacing.setLine("360");//此处1.5倍行间距
//        pSpacing.setBeforeLines(new BigInteger("20"));//段前0.2
//        pSpacing.setAfterLines(new BigInteger("10"));//段后0.1
//        //设置字体
//        CTRPr pRpr = run.getRPr();
//        CTFonts fonts = pRpr.addNewRFonts();
//        fonts.setAscii("Times New Roman");
//        fonts.setEastAsia("宋体");
//        fonts.setHAnsi("宋体");
//        // 设置字体大小
//        CTHpsMeasure sz = pRpr.addNewSz();
//        sz.setVal("24");
//
//        CTHpsMeasure szCs = pRpr.addNewSzCs();
//        szCs.setVal("24");
//    }
//
//    public static void addRow(CTSdtBlock Sdt, int level, String title, int page, String bookmarkRef) {
//        CTSdtContentBlock contentBlock = Sdt.getSdtContent();
//        CTP p = contentBlock.addNewP();
//        CTPPr pPr = p.addNewPPr();
//        pPr.addNewPStyle().setVal("TOC" + level);
//        CTTabs tabs = pPr.addNewTabs();//Set of Custom Tab Stops自定义制表符集合
//        CTTabStop tab = tabs.addNewTab();//Custom Tab Stop自定义制表符
//        tab.setVal(STTabJc.RIGHT);
//        tab.setLeader(STTabTlc.DOT);
//        tab.setPos(new BigInteger("8200"));//默认为8290，因为调整过页边距，所有需要调整，手动设置找出最佳值
//        pPr.addNewRPr().addNewNoProof();//不检查语法
//        CTR run = p.addNewR();
//        run.addNewRPr().addNewNoProof();
//        run.addNewT().setStringValue(title);//添加标题文字
//        //设置标题字体
//        CTRPr pRpr = run.getRPr();
//        CTFonts fonts = pRpr.addNewRFonts();
//        fonts.setAscii("Times New Roman");
//        fonts.setEastAsia("宋体");
//        fonts.setHAnsi("宋体");
//        // 设置标题字体大小
//        CTHpsMeasure sz = pRpr.addNewSz();
//        sz.setVal("24");
//        CTHpsMeasure szCs = pRpr.addNewSzCs();
//        szCs.setVal("24");
//        //添加制表符
//        run = p.addNewR();
//        run.addNewRPr().addNewNoProof();
//        run.addNewTab();
//        //添加页码左括号
//        p.addNewR().addNewT().setStringValue("(");
//        //STFldCharType.BEGIN标识与结尾处STFldCharType.END相对应
//        run = p.addNewR();
//        run.addNewRPr().addNewNoProof();
//        run.addNewFldChar().setFldCharType(STFldCharType.BEGIN);//Field Character Type
//        // pageref run
//        run = p.addNewR();
//        run.addNewRPr().addNewNoProof();
//        CTText text = run.addNewInstrText();//Field Code 添加域代码文本控件
//        text.setSpace(SpaceAttribute.Space.PRESERVE);
//        // bookmark reference
//        //源码的域名为" PAGEREF _Toc","\h"含义为在目录内建立目录项与页码的超链接
//        text.setStringValue(" PAGEREF " + bookmarkRef + " \\h ");
//        p.addNewR().addNewRPr().addNewNoProof();
//        run = p.addNewR();
//        run.addNewRPr().addNewNoProof();
//        run.addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
//        // page number run
//        run = p.addNewR();
//        run.addNewRPr().addNewNoProof();
//        run.addNewT().setStringValue(Integer.toString(page));
//        run = p.addNewR();
//        run.addNewRPr().addNewNoProof();
//        //STFldCharType.END标识与上面STFldCharType.BEGIN相对应
//        run.addNewFldChar().setFldCharType(STFldCharType.END);
//        //添加页码右括号
//        p.addNewR().addNewT().setStringValue(")");
//        //设置行间距
//        CTSpacing pSpacing = pPr.getSpacing() != null ? pPr.getSpacing() : pPr.addNewSpacing();
//        pSpacing.setLineRule(STLineSpacingRule.AUTO);//行间距类型：多倍
//        pSpacing.setLine(new BigInteger("360"));//此处1.5倍行间距
//    }

    public static void addTOC(PierDocument document) {
        CTSdtBlock sdt = document.getSdt();
        // sdtPr
        CTSdtPr sdtPr = sdt.addNewSdtPr();
        CTRPr rPr = sdtPr.addNewRPr();
        CTFonts rFonts = rPr.addNewRFonts();
        rFonts.setAscii("宋体");
        rFonts.setEastAsia("宋体");
        rFonts.setHAnsi("宋体");
        rPr.addNewB();
        rPr.addNewBCs();
        CTSdtDocPart docPartGallery = sdtPr.addNewDocPartObj();
        docPartGallery.addNewDocPartGallery().setVal("Table of Contents");
        docPartGallery.addNewDocPartUnique();

        CTSdtContentBlock sdtContent = sdt.addNewSdtContent();
        addMidTitle(sdtContent);

        ArrayList<PierParagraph> paragraphs = document.getParagraphs();
        for (PierParagraph par : paragraphs) {
            String parStyle = par.getStyleID();
            if (parStyle != null && parStyle.startsWith("title")) {
                try {
                    int level = Integer.parseInt(parStyle.substring("title".length()));
                    if (level == 1) {
                        //添加栏目
                        addRowOnlyTitle(sdtContent, level, par.getParagraph().getText());
                    } else if (level == 2) {
                        //添加标题
                        addRow(sdtContent, level, par.getParagraph().getText());
                    } else {
                        addRow(sdtContent, level, par.getParagraph().getText());
                    }
                } catch (NumberFormatException e) {
                    e.printStackTrace();
                }
            }
        }

        addEnd(sdtContent);

    }

    public static void addRowOnlyTitle(CTSdtContentBlock sdtContent, int level, String title) {
        CTP p = sdtContent.addNewP();
        CTPPr pPr = p.addNewPPr();

        pPr.addNewPStyle().setVal("TOC" + level);

        CTTabs tabs = pPr.addNewTabs();
        CTTabStop tab = tabs.addNewTab();
        tab.setVal(STTabJc.Enum.forString("right"));
        tab.setLeader(STTabTlc.Enum.forString("dot"));
        tab.setPos("8296");

        CTParaRPr rPr = pPr.addNewRPr();
        CTFonts rFonts = rPr.addNewRFonts();
        rFonts.setAscii("宋体");
        rFonts.setEastAsia("宋体");
        rFonts.setHAnsi("宋体");
        rPr.addNewSz().setVal("24");
        rPr.addNewSzCs().setVal("24");


        CTR r = p.addNewR();
        r.addNewFldChar().setFldCharType(STFldCharType.Enum.forString("begin"));
        CTText instrText = r.addNewInstrText();
        instrText.setStringValue(" TOC \\o \"1-3\" \\h \\z \\u ");
        instrText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));

        r.addNewFldChar().setFldCharType(STFldCharType.Enum.forString("separate"));

        CTHyperlink hyperlink = p.addNewHyperlink();
        r = hyperlink.addNewR();
        CTRPr ctrPr = r.addNewRPr();
        ctrPr.addNewRStyle().setVal("Hyperlink");
        ctrPr.addNewNoProof();
        ctrPr.addNewWebHidden();
        rFonts = ctrPr.addNewRFonts();
        rFonts.setAscii("宋体");
        rFonts.setEastAsia("宋体");
        rFonts.setHAnsi("宋体");
        ctrPr.addNewSz().setVal("24");
        ctrPr.addNewSzCs().setVal("24");

        CTText t = r.addNewT();
        t.setStringValue("1 工程概况");
        t.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
        r.addNewTab();
        r.addNewFldChar().setFldCharType(STFldCharType.Enum.forString("begin"));
        instrText = r.addNewInstrText();
        instrText.setStringValue(" PAGEREF _Toc137214757 \\h ");
        instrText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
        r.addNewFldChar().setFldCharType(STFldCharType.Enum.forString("separate"));
        r.addNewT().setStringValue("2");
        r.addNewFldChar().setFldCharType(STFldCharType.Enum.forString("end"));

    }

    public static void addRow(CTSdtContentBlock sdtContent, int level, String title) {

    }

    public static void addMidTitle(CTSdtContentBlock sdtContent) {
        CTP p = sdtContent.addNewP();
        CTPPr pPr = p.addNewPPr();
        pPr.addNewJc().setVal(STJc.Enum.forString("center"));
        CTParaRPr rPr = pPr.addNewRPr();
        CTFonts rFonts = rPr.addNewRFonts();
        rFonts.setAscii("宋体");
        rFonts.setEastAsia("宋体");
        rFonts.setHAnsi("宋体");
        rPr.addNewB();
        rPr.addNewBCs();
        rPr.addNewSz().setVal("32");
        rPr.addNewSzCs().setVal("32");

        CTR r = p.addNewR();
        r.addNewT().setStringValue("目    录");
        CTRPr ctrPr = r.addNewRPr();
        rFonts = ctrPr.addNewRFonts();
        rFonts.setAscii("宋体");
        rFonts.setEastAsia("宋体");
        rFonts.setHAnsi("宋体");
        ctrPr.addNewB();
        ctrPr.addNewBCs();
        ctrPr.addNewSz().setVal("32");
        ctrPr.addNewSzCs().setVal("32");
    }

    public static void addEnd(CTSdtContentBlock sdtContent) {
        CTP p = sdtContent.addNewP();
        CTR r = p.addNewR();
        CTRPr rPr = r.addNewRPr();
        r.addNewFldChar().setFldCharType(STFldCharType.Enum.forString("end"));
    }



}
