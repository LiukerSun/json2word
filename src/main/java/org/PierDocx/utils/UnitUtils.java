package org.PierDocx.utils;

import java.util.Collections;
import java.util.List;

import org.apache.poi.util.Units;

/**
 * @author Sayi
 */
public final class UnitUtils {

    /**
     * cm to twips
     * 
     * @param cm
     * @return in twentieths of a point (1/1440 of an inch)
     */
    public static int cm2Twips(double cm) {
        return (int) (cm / 2.54 * 1440);
    }

    /**
     * point to twips
     * 
     * @param pt
     * @return in twentieths of a point (1/1440 of an inch)
     */
    public static int point2Twips(double pt) {
        return (int) (pt * 20);
    }

    /**
     * twips to point
     * 
     * @param twips
     * @return
     */
    public static double twips2Point(int twips) {
        return (twips / 20.0);
    }

    /**
     * cm to pixel
     * 
     * @param cm
     * @return pixel
     */
    public static int cm2Pixel(double cm) {
        return Units.pointsToPixel(cm / 2.54 * 1440 / 20.0);
    }

    /**
     * twips to pixel
     * 
     * @param twips
     * @return pixel
     */
    public static int twips2Pixel(int twips) {
        return Units.pointsToPixel(twips / 20);
    }

    /**
     * average the width
     * 
     * @param width
     * @param col
     * @return
     */
    public static int[] average(int width, int col) {
        int colVal = (Integer.valueOf(width)) / col;
        List<Integer> nCopies = Collections.nCopies(col, colVal);
        return nCopies.stream().mapToInt(i -> i).toArray();
    }

}
