package org.PierDocx;

import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.util.ArrayList;


public class PierTable {
    public XWPFTable table;
    ArrayList<PierTableRow> rows_list = new ArrayList<>();


    public PierTable(PierDocument document, int rows, int cols) {
        super();
        this.table = document.document.createTable(rows, cols);
        this.table.setWidthType(TableWidthType.PCT);
        table.setWidth("100%");
        this.get_rows();
    }

    public PierTable mergeCellsHorizontal(Integer row, Integer fromCell, Integer toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = this.table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
        return this;
    }


    public PierTable mergeCellsVertically(Integer col, Integer fromRow, Integer toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = this.table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
        return this;
    }

    public ArrayList<PierTableRow> get_rows() {
        // 防止重复添加。先清空。
        if (this.rows_list != null) {
            this.rows_list.clear();
        }
        for (XWPFTableRow row : this.table.getRows()) {
            this.rows_list.add(new PierTableRow(row));
        }
        return this.rows_list;
    }

    public PierTableRow get_row(int row_index) {
        return this.rows_list.get(row_index);
    }


    public static class PierTableRow {
        XWPFTableRow row;
        ArrayList<PierTableCell> cells = new ArrayList<>();

        public PierTableRow(XWPFTableRow row) {
            this.row = row;
            this.get_cells();
        }

        public ArrayList<PierTableCell> get_cells() {
            // 防止重复添加。先清空。
            if (this.cells != null) {
                this.cells.clear();
            }
            for (XWPFTableCell cell : this.row.getTableCells()) {
                this.cells.add(new PierTableCell(cell));
            }
            return this.cells;
        }

        public PierTableCell get_cell(int cell_index) {
            return this.cells.get(cell_index);
        }


    }


    public static class PierTableCell {
        XWPFTableCell cell;
        ArrayList<PierParagraph> paragraphs = new ArrayList<>();
        int paragraphs_count = 0;


        public PierTableCell(XWPFTableCell cell) {
            this.cell = cell;
        }

        public PierTableCell setText(String text) {
            this.cell.setText(text);
            return this;
        }

        public PierTableCell setWidth(String width) {
            this.cell.setWidth(width);
            return this;
        }

        public PierTableCell setVerticalAlignment(XWPFTableCell.XWPFVertAlign VertAlign) {
            this.cell.setVerticalAlignment(VertAlign);
            return this;
        }

        public PierParagraph addParagraph() {
            PierParagraph paragraph = new PierParagraph(this);
            this.paragraphs.add(paragraph);
            this.paragraphs_count += 1;
            return paragraph;
        }


    }


}


