package org.PierDocx;

import com.fasterxml.jackson.databind.JsonNode;
import org.PierDocx.style.TableStyle;
import org.PierDocx.utils.Preconditions;
import org.PierDocx.utils.StyleUtils;
import org.PierDocx.utils.TableTools;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.ArrayList;

import static org.PierDocx.utils.TableTools.*;
import static org.test.Main.logger;


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

    public PierTable addAllBorder() {
        TableTools.setAllBorder(this);
        return this;
    }

    private void mergeCellsHorizontal(int row, int fromCell, int toCell) {
        TableTools.mergeCellsHorizonal(this, row, fromCell, toCell);
    }

    private void mergeCellsVertically(int col, int fromRow, int toRow) {
        TableTools.mergeCellsVertically(this, col, fromRow, toRow);
    }

    public void mergeCell(int firstRow, int firstColumn, int lastRow, int lastColumn) {
        Preconditions.requireDiffCell(firstRow, firstColumn, lastRow, lastColumn, "Need Different Cells!");
        Preconditions.requireBiggerThan(lastColumn, firstColumn, "lastColumn need bigger than firstColumn!");
        Preconditions.requireBiggerThan(lastRow, firstRow, "lastRow need bigger than firstRow!");

        String _mergeType = "Both";
        if (firstColumn == lastColumn) {
            _mergeType = "ColOnly";
        }
        if (firstRow == lastRow) {
            _mergeType = "RowOnly";
        }


        switch (_mergeType) {
            case "Both": {
                for (int row = firstRow; row <= lastRow; row++) {
                    this.mergeCellsHorizontal(row, firstColumn, lastColumn);
                }
                this.mergeCellsVertically(firstColumn, firstRow, lastRow);
                break;
            }
            case "RowOnly": {
                this.mergeCellsHorizontal(firstRow, firstColumn, lastColumn);
                break;
            }
            case "ColOnly": {
                this.mergeCellsVertically(firstColumn, firstRow, lastRow);
                break;
            }
        }


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

    private int getRowsSize() {
        return obtainRowSize(this);
    }

    private int getColsSize() {
        return obtainColumnSize(this);
    }


}


