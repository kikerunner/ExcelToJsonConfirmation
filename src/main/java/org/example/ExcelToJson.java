package org.example;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

public class ExcelToJson {
    public static void main(String[] args) {
        String folderPath = "/home/kike/ExcelToJsonConfirmation/0Input";  // Cambia esto a la ruta de tu carpeta
        String outputFolderPath = "/home/kike/ExcelToJsonConfirmation/1Output";  // Cambia esto a la ruta de tu carpeta de salida
        processAllFilesInFolder(folderPath, outputFolderPath);
    }

    private static void processAllFilesInFolder(String folderPath, String outputFolderPath) {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(folderPath), "*.xlsx")) {
            for (Path entry : stream) {
                List<Map<String, String>> data = readExcel(entry.toString());
                String jsonArray = createJsonString(data);
                String outputFilePath = generateOutputFilePath(entry, outputFolderPath);
                saveJsonToFile(jsonArray, outputFilePath);
                System.out.println("JSON guardado en: " + outputFilePath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<Map<String, String>> readExcel(String filePath) {
        List<Map<String, String>> data = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            Row headerRow = rowIterator.next();
            List<String> headers = new ArrayList<>();
            headerRow.forEach(cell -> headers.add(cell.getStringCellValue()));

            while (rowIterator.hasNext()) {
                Row currentRow = rowIterator.next();
                Map<String, String> map = new HashMap<>();
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = currentRow.getCell(i);
                    String cellValue = cell != null ? cell.toString() : "";
                    map.put(headers.get(i), cellValue);
                }
                data.add(map);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return data;
    }

    public static String createJsonString(List<Map<String, String>> data) {
        StringBuilder jsonStringBuilder = new StringBuilder();
        jsonStringBuilder.append("[");

        for (Map<String, String> record : data) {
            // Iniciar el objeto JSON
            jsonStringBuilder.append("{");

            // Añadir el objeto "missingAttributeEventInfo"
            jsonStringBuilder.append("\"missingAttributeEventInfo\": {");

            // Añadir los campos en el orden deseado
            jsonStringBuilder.append("\"ouid\":\"").append(record.get("ouid")).append("\",");  // 1
            jsonStringBuilder.append("\"timeStamp\":\"").append(getCurrentTimestamp()).append("\",");  // 2
            jsonStringBuilder.append("\"deliveryDate\":\"").append(formatDeliveryDate(record.get("deliveryDate"))).append("\",");  // 3
            jsonStringBuilder.append("\"pbkCode\":\"Kleider\",");  // 4
            jsonStringBuilder.append("\"missingAttributesList\":[],");  // 5
            jsonStringBuilder.append("\"supplier\":\"").append(record.get("supplier").split("\\.")[0]).append("\"");  // 6

            // Cerrar el objeto "missingAttributeEventInfo"
            jsonStringBuilder.append("}");  // Cierra el objeto

            // Cerrar el objeto JSON principal
            jsonStringBuilder.append("},");

        }

        // Eliminar la última coma y cerrar el array
        if (jsonStringBuilder.length() > 1) {
            jsonStringBuilder.setLength(jsonStringBuilder.length() - 1);  // Eliminar la última coma
        }
        jsonStringBuilder.append("]");

        return jsonStringBuilder.toString();
    }

    // Función para obtener el timestamp actual en el formato ISO 8601 con 'Z'
    private static String getCurrentTimestamp() {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
        sdf.setTimeZone(TimeZone.getTimeZone("UTC"));  // Aseguramos que esté en UTC
        return sdf.format(new Date());
    }

    // Función para formatear la deliveryDate a 'yyyy-MM-dd' (solo la fecha)
    private static String formatDeliveryDate(String deliveryDate) {
        // Suponemos que el formato original incluye más información que queremos eliminar
        return deliveryDate.split(" ")[0];  // Tomamos solo la parte de la fecha (antes del espacio)
    }


    private static void saveJsonToFile(String jsonArray, String filePath) {
        try (FileWriter file = new FileWriter(filePath)) {
            file.write(jsonArray);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String generateOutputFilePath(Path inputFile, String outputFolderPath) {
        String fileName = inputFile.getFileName().toString();
        String baseName = fileName.substring(0, fileName.lastIndexOf('.'));
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        return Paths.get(outputFolderPath, baseName + "_" + timestamp + ".txt").toString();
    }
}