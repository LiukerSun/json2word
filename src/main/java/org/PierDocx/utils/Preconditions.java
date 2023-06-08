package org.PierDocx.utils;

public final class Preconditions {

    private Preconditions() {
    }

    public static void requireGreaterThan(int first, int second, String message) {
        if (first <= second) {
            throw new IllegalStateException(message);
        }
    }

    public static void requireBiggerThan(int first, int second, String message) {
        if (first < second) {
            throw new IllegalStateException(message);
        }
    }

    public static void requireDiffCell(int firstRow, int firstColumn, int lastRow, int lastColumn, String message) {
        if ((firstRow == lastRow) & (firstColumn == lastColumn)) {
            throw new IllegalStateException(message);
        }
    }
}
