package org.PierDocx;

import org.PierDocx.style.TableStyle;
import org.PierDocx.utils.StyleUtils;
import org.PierDocx.utils.TableTools;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.ArrayList;

import static org.PierDocx.utils.TableTools.*;


public class PierTable {
    public XWPFTable table;
    ArrayList<PierTableRow> rows_list = new ArrayList<>();


    public PierTable(PierDocument document, int rows, int cols) {
        super();
        this.table = document.document.createTable(rows, cols);
        initBasicTable(this, cols, 0);
        this.table.setWidthType(TableWidthType.PCT);
        table.setWidth("100%");
        this.getRows();
    }

    public PierTable addStyle(PierTable table, TableStyle style) {
        StyleUtils.styleTable(table, style);
        return table;
    }

    public PierTable addStyle(TableStyle style) {
        StyleUtils.styleTable(this, style);
        return this;
    }

    public PierTable addAllBorder(){
        TableTools.setAllBorder(this);
        return this;
    }

    public void mergeCellsHorizontal(Integer row, Integer fromCell, Integer toCell) {
        TableTools.mergeCellsHorizonal(this, row, fromCell, toCell);
    }

    public void mergeCellsVertically(Integer col, Integer fromRow, Integer toRow) {
        TableTools.mergeCellsVertically(this, col, fromRow, toRow);
    }

    public ArrayList<PierTableRow> getRows() {
        // 防止重复添加。先清空。
        if (this.rows_list != null) {
            this.rows_list.clear();
        }
        for (XWPFTableRow row : this.table.getRows()) {
            this.rows_list.add(new PierTableRow(row));
        }
        return this.rows_list;
    }

    public PierTableRow getRow(int row_index) {
        return this.rows_list.get(row_index);
    }

    public int getRowsSize() {
        return obtainRowSize(this);
    }

    public int getColsSize() {
        return obtainColumnSize(this);
    }


}


