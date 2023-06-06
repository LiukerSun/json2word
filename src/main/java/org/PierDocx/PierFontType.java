package org.PierDocx;

import java.util.HashMap;
import java.util.Map;

public enum PierFontType {
    ENG(1),
    CN(2);
    private final int value;

    public int getValue() {
        return this.value;
    }

    private static final Map<Integer, PierFontType> imap = new HashMap<>();

    public static PierFontType valueOf(int type) {
        PierFontType err = (PierFontType) imap.get(type);
        if (err == null) {
            throw new IllegalArgumentException("Unknown Pier Font Type: " + type);
        } else {
            return err;
        }
    }

    private PierFontType(int val) {
        this.value = val;
    }

    static {
        PierFontType[] var0 = values();
        int var1 = var0.length;

        for (PierFontType p : var0) {
            imap.put(p.getValue(), p);
        }

    }
}
