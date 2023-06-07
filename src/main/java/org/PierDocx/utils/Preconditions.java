package org.PierDocx.utils;

public final class Preconditions {

    private Preconditions() {
    }

    public static void requireGreaterThan(int first, int second, String message) {
        if (first <= second) {
            throw new IllegalStateException(message);
        }
    }
}
