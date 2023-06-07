package org.PierDocx.utils;

import org.PierDocx.PierTable;
import org.PierDocx.PierTableCell;
import org.PierDocx.PierTableRow;
import org.PierDocx.style.BorderStyle;
import org.PierDocx.style.CellStyle;
import org.PierDocx.style.RowStyle;
import org.PierDocx.style.TableStyle;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;

public final class TableTools {

    public static void mergeCellsHorizonal(PierTable table, int row, int fromCol, int toCol) {
        Preconditions.requireGreaterThan(toCol, fromCol, "To column to be merged must greater than from column.");
        mergeCellsHorizontalWithoutRemove(table, row, fromCol, toCol);
        PierTableRow rowTable = table.getRow(row);
        for (int colIndex = fromCol + 1; colIndex <= toCol; colIndex++) {
            rowTable.row.removeCell(fromCol + 1);
            if (rowTable.row.getTableCells().size() != rowTable.row.getCtRow().sizeOfTcArray()) {
                rowTable.row.getCtRow().removeTc(fromCol + 1);
            }
        }
    }

    public static void mergeCellsHorizontalWithoutRemove(PierTable table, int row, int fromCol, int toCol) {
        Preconditions.requireGreaterThan(toCol, fromCol, "To column to be merged must greater than from column.");
        PierTableCell cell = table.getRow(row).getCell(fromCol);
        CTTcPr tcPr = getTcPr(cell.cell);
        tcPr.addNewGridSpan();
        tcPr.getGridSpan().setVal(BigInteger.valueOf((long) (toCol - fromCol + 1)));
        int tcw = 0;
        for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
            PierTableCell tableCell = table.getRow(row).getCell(colIndex);
            // TODO pct, auto
            if (TableWidthType.DXA == tableCell.cell.getWidthType()) {
                if (-1 == tableCell.cell.getWidth()) return;
                tcw += tableCell.cell.getWidth();
            } else {
                return;
            }
        }
        if (0 != tcw) cell.setWidth(tcw + "");
    }

    public static void mergeCellsVertically(PierTable table, int col, int fromRow, int toRow) {
        Preconditions.requireGreaterThan(toRow, fromRow, "To row to be merged must greater than from row.");
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.table.getRow(rowIndex).getCell(col);
            CTTcPr tcPr = getTcPr(cell);
            CTVMerge vMerge = tcPr.addNewVMerge();
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                vMerge.setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                vMerge.setVal(STMerge.CONTINUE);
            }
        }
    }

    private static void ensureTblW(PierTable table) {
        CTTbl ctTbl = table.table.getCTTbl();
        CTTblPr tblPr = (ctTbl.getTblPr() != null) ? ctTbl.getTblPr() : ctTbl.addNewTblPr();
        if (!tblPr.isSetTblW()) tblPr.addNewTblW();
    }

    public static void borderTable(PierTable table, int size) {
        CTTblPr tblPr = getTblPr(table);
        CTTblBorders tblBorders = tblPr.getTblBorders();
        if (null == tblBorders) {
            tblBorders = tblPr.addNewTblBorders();
        }
        BigInteger borderSize = BigInteger.valueOf(size);
        if (!tblBorders.isSetBottom()) tblBorders.addNewBottom();
        if (!tblBorders.isSetLeft()) tblBorders.addNewLeft();
        if (!tblBorders.isSetTop()) tblBorders.addNewTop();
        if (!tblBorders.isSetRight()) tblBorders.addNewRight();
        if (!tblBorders.isSetInsideH()) tblBorders.addNewInsideH();
        if (!tblBorders.isSetInsideV()) tblBorders.addNewInsideV();
        tblBorders.getBottom().setSz(borderSize);
        tblBorders.getLeft().setSz(borderSize);
        tblBorders.getTop().setSz(borderSize);
        tblBorders.getRight().setSz(borderSize);
        tblBorders.getInsideH().setSz(borderSize);
        tblBorders.getInsideV().setSz(borderSize);
    }

    public static void initBasicTable(PierTable table, int col, float width) {
        int defaultBorderSize = 4;
        widthTable(table, width, col);
        borderTable(table, defaultBorderSize);
//        styleTable(table, style);
    }

    public static boolean isInsideTable(XWPFRun run) {
        return ((XWPFParagraph) run.getParent()).getPartType() == BodyType.TABLECELL;
    }

    public static void styleTable(PierTable table, TableStyle style) {
        StyleUtils.styleTable(table, style);
    }

    public static int obtainRowSize(PierTable table) {
        return table.getRows().size();
    }

    public static int obtainColumnSize(PierTable table) {
        return table.getRows().get(0).row.getTableCells().size();
    }

    private static CTTblGrid getTblGrid(PierTable table) {
        CTTblGrid tblGrid = table.table.getCTTbl().getTblGrid();
        if (null == tblGrid) {
            tblGrid = table.table.getCTTbl().addNewTblGrid();
        }
        return tblGrid;
    }

    private static CTTblLayoutType getTblLayout(PierTable table) {
        CTTblPr tblPr = getTblPr(table);
        return tblPr.isSetTblLayout() ? tblPr.getTblLayout() : tblPr.addNewTblLayout();
    }

    private static CTTblPr getTblPr(PierTable table) {
        CTTblPr tblPr = table.table.getCTTbl().getTblPr();
        if (null == tblPr) {
            tblPr = table.table.getCTTbl().addNewTblPr();
        }
        return tblPr;
    }

    private static CTTcPr getTcPr(XWPFTableCell cell) {
        return cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
    }

    public static void setAllBorder(PierTable table) {
        TableStyle tableStyle = new TableStyle();
        tableStyle.setLeftBorder(new BorderStyle().setDefaultBorderStyle());
        tableStyle.setRightBorder(new BorderStyle().setDefaultBorderStyle());
        tableStyle.setTopBorder(new BorderStyle().setDefaultBorderStyle());
        tableStyle.setBottomBorder(new BorderStyle().setDefaultBorderStyle());
        tableStyle.setInsideHBorder(new BorderStyle().setDefaultBorderStyle());
        tableStyle.setInsideVBorder(new BorderStyle().setDefaultBorderStyle());
        styleTable(table,tableStyle);
    }

    @Deprecated
    public static void widthTable(PierTable table, float[] colWidths) {
        float widthCM = 0;
        for (float w : colWidths) {
            widthCM += w;
        }
        int width = UnitUtils.cm2Twips(widthCM);
        CTTblPr tblPr = getTblPr(table);
        CTTblWidth tblW = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        tblW.setType(0 == width ? STTblWidth.AUTO : STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(width));

        if (0 != width) {
            CTTblGrid tblGrid = getTblGrid(table);
            for (float w : colWidths) {
                CTTblGridCol addNewGridCol = tblGrid.addNewGridCol();
                addNewGridCol.setW(BigInteger.valueOf(UnitUtils.cm2Twips(w)));
            }
        }
    }

    @Deprecated
    public static void widthTable(PierTable table, float widthCM, int cols) {
        int width = UnitUtils.cm2Twips(widthCM);
        CTTblPr tblPr = getTblPr(table);
        CTTblWidth tblW = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        tblW.setType(0 == width ? STTblWidth.AUTO : STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(width));

        if (0 != width) {
            CTTblGrid tblGrid = getTblGrid(table);
            for (int j = 0; j < cols; j++) {
                CTTblGridCol addNewGridCol = tblGrid.addNewGridCol();
                addNewGridCol.setW(BigInteger.valueOf(width / cols));
            }
        }
    }

}
