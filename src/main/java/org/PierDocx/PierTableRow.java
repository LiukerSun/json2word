package org.PierDocx;

import org.PierDocx.style.RowStyle;
import org.PierDocx.utils.StyleUtils;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.ArrayList;

public class PierTableRow {
    public XWPFTableRow row;
    ArrayList<PierTableCell> cells = new ArrayList<>();

    public PierTableRow(XWPFTableRow row) {
        this.row = row;
        this.getCells();
    }

    public PierTableRow addStyle(PierTableRow row, RowStyle style) {
        StyleUtils.styleTableRow(row, style);
        return row;
    }

    public PierTableRow addStyle(RowStyle style) {
        StyleUtils.styleTableRow(this, style);
        return this;
    }


    public ArrayList<PierTableCell> getCells() {
        // 防止重复添加。先清空。
        if (this.cells != null) {
            this.cells.clear();
        }
        for (XWPFTableCell cell : this.row.getTableCells()) {
            this.cells.add(new PierTableCell(cell));
        }
        return this.cells;
    }



    public PierTableCell getCell(int cell_index) {
        return this.cells.get(cell_index);
    }


}
