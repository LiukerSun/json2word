package org.PierDocx.utils;

import java.util.HashMap;

public class PicUtils {
    public static int get_pic_type(String pic_path) {
        String[] tmp = pic_path.split("\\.");
        String pic_type = tmp[tmp.length - 1].toUpperCase();
        HashMap<String, Integer> type_map = new HashMap<>() {};
        type_map.put("EMF",2);
        type_map.put("WMF",3);
        type_map.put("PICT",4);
        type_map.put("JPG",5);
        type_map.put("JPEG",5);
        type_map.put("PNG",6);
        type_map.put("DIB",7);
        type_map.put("GIF",8);
        type_map.put("TIFF",9);
        type_map.put("EPS",10);
        type_map.put("BMP",11);
        type_map.put("WPG",12);
        return type_map.getOrDefault(pic_type, 1);
    }
}
