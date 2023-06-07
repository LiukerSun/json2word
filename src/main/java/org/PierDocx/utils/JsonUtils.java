package org.PierDocx.utils;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.FileInputStream;
import java.io.InputStream;

import static org.test.Main.logger;

public class JsonUtils {
    public static JsonNode loadJson(String json_path) throws JsonProcessingException {
        StringBuilder sb = new StringBuilder();
        try (InputStream input = new FileInputStream(json_path)) {
            byte[] buffer = new byte[1024];
            int length;
            length = input.read(buffer);
            while (length != -1) {
                sb.append(new String(buffer, 0, length));
                length = input.read(buffer);
            }
        } catch (Exception e) {
            logger.error("load json fail.");
            e.printStackTrace();
        }
        ObjectMapper mapper = new ObjectMapper();
        return mapper.readTree(sb.toString());
    }

}
