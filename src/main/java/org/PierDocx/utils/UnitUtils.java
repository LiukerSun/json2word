package org.PierDocx.utils;

import java.util.Collections;
import java.util.List;

import org.apache.poi.util.Units;


public final class UnitUtils {

    public static int cm2Twips(double cm) {
        return (int) (cm / 2.54 * 1440);
    }

    public static int point2Twips(double pt) {
        return (int) (pt * 20);
    }

    public static double twips2Point(int twips) {
        return (twips / 20.0);
    }

    public static int cm2Pixel(double cm) {
        return Units.pointsToPixel(cm / 2.54 * 1440 / 20.0);
    }

    public static int twips2Pixel(int twips) {
        return Units.pointsToPixel(twips / 20);
    }

    public static int[] average(int width, int col) {
        int colVal = (Integer.valueOf(width)) / col;
        List<Integer> nCopies = Collections.nCopies(col, colVal);
        return nCopies.stream().mapToInt(i -> i).toArray();
    }

}
