package utils;

import java.io.Closeable;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils implements Closeable {
    private Workbook workbook;
    private Sheet sheet;
    private int count = 0;
    private int totalCount = 0;

    public ExcelUtils(String excelPath, String sheetName) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(excelPath)) {
            workbook = new XSSFWorkbook(fileInputStream);
            sheet = workbook.getSheet(sheetName);
        }
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void concatenateEFGAndStoreInColumnK() {
        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);

            // Add a null check for the row object
            if (row == null) {
                continue; // Skip this row if it's null
            }

            // Concatenate cell E, F, G, and store the result in column K
            Cell cellE = row.getCell(4);
            Cell cellF = row.getCell(5);
            Cell cellG = row.getCell(6);

            String concatenatedValue = getCellStringValue(cellE) + getCellStringValue(cellF) + getCellStringValue(cellG);

            Cell kCell = row.createCell(10);
            kCell.setCellValue(concatenatedValue);

            // Copy the cell value of K to column L
            Cell lCell = row.createCell(11);
            lCell.setCellValue(kCell.getStringCellValue());
        }
    }

    public void processColumnK() {
        List<String> seenValues = new ArrayList<>();

        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);

            // Add a null check for the row object
            if (row == null) {
                continue; // Skip this row if it's null
            }

            Cell kCell = row.getCell(10);

            // Add a null check for the kCell
            if (kCell != null) {
                String kCellValue = getCellStringValue(kCell);

                if (kCellValue == null || kCellValue.trim().isEmpty()) {
                    // Skip the row if K cell is blank, empty, or null
                    continue;
                }

                if (shouldSkip(kCellValue)) {
                    seenValues.add(kCellValue);
                    continue;
                }

                int occurrenceCount = countOccurrences(kCellValue, seenValues);

                if (occurrenceCount > 0) {
                    // Check if column J is empty or null
                    Cell jCell = row.getCell(9);

                    // Add a null check for the jCell
                    if (jCell != null) {
                        String jCellValue = getCellStringValue(jCell);

                        if (jCellValue == null || jCellValue.isEmpty()) {
                            // Delete the cell value of column G and write '0' in column H
                            Cell gCell = row.getCell(6);
                            if (gCell != null) {
                                gCell.setCellValue(""); // Delete the value in column G cell
                            }

                            Cell hCell = row.getCell(7);
                            if (hCell != null) {
                                hCell.setCellValue(0);
                            } else {
                                hCell = row.createCell(7);
                                hCell.setCellValue(0);
                            }

                            // Set "REPEAT" in column J
                            jCell.setCellValue("REPEAT");
                        }
                    }
                }
            }
        }
    }

    public void processColumnG() {
        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);

            // Add a null check for the row object
            if (row == null) {
                continue; // Skip this row if it's null
            }

            Cell kCell = row.getCell(10);
            Cell gCell = row.getCell(6);

            // Add a null check for the kCell and gCell
            if (kCell != null && gCell != null) {
                String kCellValue = getCellStringValue(kCell).toLowerCase();

                if (!shouldSkip(kCellValue)) { // Check if it's not a predefined skip word
                    String gCellValue = getCellStringValue(gCell);
                    if (gCellValue != null && !gCellValue.isEmpty() && !gCellValue.equalsIgnoreCase("null")) {
                        try {
                            int value = Integer.parseInt(gCellValue);
                            totalCount += value; // Add the value to totalCount
                        } catch (NumberFormatException e) {
                            // Handle non-numeric values if necessary
                        }
                    }
                }
            }
        }
    }

    public void saveExcel(String outputPath) throws IOException {
        try (FileOutputStream fileOutputStream = new FileOutputStream(outputPath)) {
            workbook.write(fileOutputStream);
        }
    }

    public String getCellStringValue(Cell cell) {
        DataFormatter dataFormatter = new DataFormatter();
        return dataFormatter.formatCellValue(cell);
    }

    public boolean shouldSkip(String value) {
        String[] skipWords = {"Name", "Date", "Case", "Type", "Total", "No"};
        for (String word : skipWords) {
            if (value.startsWith(word)) {
                return true;
            }
        }
        return false;
    }

    private int countOccurrences(String value, List<String> seenValues) {
        int count = 0;
        for (String seenValue : seenValues) {
            if (seenValue.equals(value)) {
                count++;
            }
        }
        seenValues.add(value);
        return count;
    }

    // Close method to be used in try-with-resources when needed
    @Override
    public void close() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }
}
