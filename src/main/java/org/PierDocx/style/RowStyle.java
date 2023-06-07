package org.PierDocx.style;

import java.io.Serializable;


public class RowStyle implements Serializable {
    private int height;
    private String heightRule;
    private boolean breakAcrossPage = true;
    private boolean repeated;
    private CellStyle defaultCellStyle;
    // setter
    public void setHeight(int height) {
        this.height = height;
    }

    public void setHeightRule(String heightRule) {
        this.heightRule = heightRule;
    }

    public void setBreakAcrossPage(boolean breakAcrossPage) {
        this.breakAcrossPage = breakAcrossPage;
    }

    public void setRepeated(boolean repeated) {
        this.repeated = repeated;
    }

    public void setDefaultCellStyle(CellStyle defaultCellStyle) {
        this.defaultCellStyle = defaultCellStyle;
    }

    // getter
    public int getHeight() {
        return height;
    }

    public String getHeightRule() {
        return heightRule;
    }

    public boolean isBreakAcrossPage() {
        return breakAcrossPage;
    }

    public boolean isRepeated() {
        return repeated;
    }

    public CellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }

}
