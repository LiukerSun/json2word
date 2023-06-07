package org.PierDocx.style;

import java.util.HashMap;
import java.util.Map;

public enum XWPFHighlightColor {
    // STHighlightColor
    BLACK(1),
    BLUE(2),
    CYAN(3),
    GREEN(4),
    MAGENTA(5),
    RED(6),
    YELLOW(7),
    WHITE(8),
    DARK_BLUE(9),
    DARK_CYAN(10),
    DARK_GREEN(11),
    DARK_MAGENTA(12),
    DARK_RED(13),
    DARK_YELLOW(14),
    DARK_GRAY(15),
    LIGHT_GRAY(16),
    NONE(17);

    private static Map<Integer, XWPFHighlightColor> imap = new HashMap<>();

    static {
        for (XWPFHighlightColor p : values()) {
            imap.put(p.getValue(), p);
        }
    }

    private final int value;

    XWPFHighlightColor(int val) {
        value = val;
    }

    public static XWPFHighlightColor valueOf(int type) {
        XWPFHighlightColor err = imap.get(type);
        if (err == null) throw new IllegalArgumentException("Unknown HighlightColor : " + type);
        return err;
    }

    public int getValue() {
        return value;
    }

}
