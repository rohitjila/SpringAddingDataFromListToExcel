package com.springcore;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ExcelWriter {

    public static void writeToExcel(List<Map<String, Object>> data, String filePath) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Employee Data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        int cellIndex = 0;
        for (String key : data.get(0).keySet()) {
            Cell cell = headerRow.createCell(cellIndex++);
            cell.setCellValue(key);
        }

        // Create data rows
        int rowIndex = 1;
        for (Map<String, Object> row : data) {
            Row dataRow = sheet.createRow(rowIndex++);
            cellIndex = 0;
            for (Object value : row.values()) {
                Cell cell = dataRow.createCell(cellIndex++);
                if (value instanceof String) {
                    cell.setCellValue((String) value);
                } else if (value instanceof Number) {
                    cell.setCellValue(((Number) value).doubleValue());
                }
            }
        }

        // Write to Excel file
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
            System.out.println("Excel file written successfully at " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}

