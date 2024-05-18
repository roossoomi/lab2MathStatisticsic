package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

public class ExportModule {
    public static void writeExcelFile(HashMap<String,HashMap<String, Double>> map,HashMap<String, Double> mapCovariance) {
        File file = new File("data.xlsx");
        if (file.exists()) {
            // Удалить файл, если он существует
            file.delete();
        }
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Data");

        // Записать заголовки столбцов
        Row headerRow = sheet.createRow(0);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue(" ");
        int colIndex = 1;
        String rt = null;

        for (String keyName : map.keySet()) {
            rt = keyName;
            break;
        }

        for (String outerKey : map.get(rt).keySet()) {
            headerCell = headerRow.createCell(colIndex++);
            headerCell.setCellValue(outerKey);
        }

        // Записать данные
        int rowIndex = 1;
        for (String outerKey : map.keySet()) {
            Row dataRow = sheet.createRow(rowIndex++);
            Cell keyCell = dataRow.createCell(0);
            keyCell.setCellValue(outerKey);

            HashMap<String, Double> innerMap = map.get(outerKey);
            colIndex = 1;
            for (String innerKey : innerMap.keySet()) {
                Cell dataCell = dataRow.createCell(colIndex++);
                dataCell.setCellValue(innerMap.get(innerKey));
            }
        }
        for (String key : mapCovariance.keySet()) {
            Row covRow = sheet.createRow(rowIndex++);
            Cell covKeyCell = covRow.createCell(0);
            covKeyCell.setCellValue(key);
            Cell covValueCell = covRow.createCell(1);
            covValueCell.setCellValue(mapCovariance.get(key));
        }

        // Сохранить лист Excel в файл

        try (FileOutputStream outputStream = new FileOutputStream("data.xlsx")) {
            workbook.write(outputStream);
            System.out.println("Файл Excel успешно создан.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}