
package org.example;

import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;

public class InputModule {
    public static HashMap<String, double[]> extractDataFromSheet(Sheet selectedSheet) {
        int columnCount = selectedSheet.getRow(0).getLastCellNum();


        HashMap<String, double[]> cellData = new HashMap<>();

        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {

            String cellName = selectedSheet.getRow(0).getCell(columnIndex).getStringCellValue();

            double[] values = new double[selectedSheet.getLastRowNum()];

            for (int rowIndex = 1; rowIndex <= selectedSheet.getLastRowNum(); rowIndex++) {
                Row row = selectedSheet.getRow(rowIndex);

                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            values[rowIndex - 1] = cell.getNumericCellValue();
                            break;
                        case FORMULA:
                            values[rowIndex - 1] = cell.getNumericCellValue();
                            break;
                        default:
                            values[rowIndex - 1] = Double.NaN;
                            break;
                    }
                } else {
                    values[rowIndex - 1] = Double.NaN;
                }
            }

            // Добавляем данные страны в HashMap
            cellData.put(cellName, values);
        }

        return cellData;
    }
}