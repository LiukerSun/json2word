package org.PierDocx.utils;

public class pic {
    public static int get_pic_type(String pic_path) {
        String[] tmp = pic_path.split("\\.");
        String pic_type = tmp[tmp.length - 1].toUpperCase();
        switch (pic_type) {
            case "EMF" -> {
                return 2;
            }
            case "WMF" -> {
                return 3;
            }
            case "PICT" -> {
                return 4;
            }
            case "JPEG", "JPG" -> {
                return 5;
            }
            case "PNG" -> {
                return 6;
            }
            case "DIB" -> {
                return 7;
            }
            case "GIF" -> {
                return 8;
            }
            case "TIFF" -> {
                return 9;
            }
            case "EPS" -> {
                return 10;
            }
            case "BMP" -> {
                return 11;
            }
            case "WPG" -> {
                return 12;
            }
            default -> {
                return 1;
            }
        }
    }
}
