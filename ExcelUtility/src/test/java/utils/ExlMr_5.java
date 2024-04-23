package utils; 
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Pattern;
import java.util.Set;
import java.util.LinkedHashSet;

public class ExlMr_5 {
    public static void main(String[] args) {
        try {
            FileInputStream fis = new FileInputStream("C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_4.xlsx");
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet("Sheet1");

            if (sheet == null) {
                System.out.println("Sheet 'Sheet1' not found.");
                return;
            }

            DataFormatter dataFormatter = new DataFormatter();

            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                if (row != null) {
                    // Delete cell values in columns K and L
                    Cell cellK = row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Cell cellL = row.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellK.setCellValue("");
                    cellL.setCellValue("");

                    // Concatenate E and F cell values into K
                    Cell cellE = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Cell cellF = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String cellEValue = dataFormatter.formatCellValue(cellE);
                    String cellFValue = dataFormatter.formatCellValue(cellF);
                    cellK.setCellValue(cellEValue + cellFValue);

                    // Check if cell K contains skip words or is empty
                    String cellKValue = dataFormatter.formatCellValue(cellK);
                    if (cellKValue != null && !cellKValue.trim().isEmpty() && !containsSkipWord(cellKValue)) {
                        // Remove the last underscore along with trailing spaces from K and store in L
                        cellKValue = cellKValue.replaceAll("\\s*_+\\s*$", "");
                        cellL.setCellValue(cellKValue);
                    }
                }
            }

            // Call the method to extract unique cell values from column L to M
            extractUniqueValuesFromKtoM(sheet);

            // Call the method to delete cell values in column M that contain skip words (individually and as substrings) and shift rows up
            deleteCellValuesContainingSkipWordsAndShiftRows(sheet);

            fis.close();

            FileOutputStream fos = new FileOutputStream("C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_5.xlsx");
            workbook.write(fos);
            fos.close();

            System.out.println("Excel manipulation completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Helper method to check if the value contains a skip word
    private static boolean containsSkipWord(String value) {
        // List of skip words (case-insensitive comparison)
        String[] skipWords = {"Name", "Date", "Case", "Type", "Total", "No"};

        for (String skipWord : skipWords) {
            // Use a regex pattern to match the skip word as a whole word (case-insensitively)
            String pattern = ".*\\b" + Pattern.quote(skipWord) + "\\b.*";
            if (Pattern.matches(pattern, value.toLowerCase())) {
                return true;
            }
        }
        return false;
    }

    // Method to delete cell values in column M that contain skip words (individually and as substrings) and shift rows up
    private static void deleteCellValuesContainingSkipWordsAndShiftRows(Sheet sheet) {
        DataFormatter dataFormatter = new DataFormatter();
        int lastRowIndex = sheet.getLastRowNum();

        for (int rowIndex = 0; rowIndex <= lastRowIndex; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cellM = row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String cellMValue = dataFormatter.formatCellValue(cellM);

                if (cellMValue != null && !cellMValue.trim().isEmpty() && containsSkipWord(cellMValue)) {
                    // Delete the cell value
                    cellM.setCellValue("");

                    // Shift rows up to fill the gap
                    for (int i = rowIndex; i < lastRowIndex; i++) {
                        Row currentRow = sheet.getRow(i);
                        Row nextRow = sheet.getRow(i + 1);

                        if (currentRow != null && nextRow != null) {
                            Cell currentCellM = currentRow.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            Cell nextCellM = nextRow.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            String nextCellMValue = dataFormatter.formatCellValue(nextCellM);

                            // Copy the value from the next row to the current row
                            currentCellM.setCellValue(nextCellMValue);

                            // Clear the value in the next row
                            nextCellM.setCellValue("");
                        }
                    }

                    // Decrement the last row index as a row is removed
                    lastRowIndex--;
                    rowIndex--; // Decrement the index to recheck the current row
                }
            }
        }
    }

    // Method to extract unique cell values from column L to M
//    private static void extractUniqueValuesFromLtoM(Sheet sheet) {
//        DataFormatter dataFormatter = new DataFormatter();
//        Set<String> uniqueValues = new LinkedHashSet<>();
//        String[] skipWords = {"Name", "Date", "Case", "Type", "Total", "No"};
//
//        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
//            Row row = sheet.getRow(rowIndex);
//
//            if (row != null) {
//                Cell cellL = row.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
//
//                String cellLValue = dataFormatter.formatCellValue(cellL);
//
//                if (cellLValue != null && !cellLValue.trim().isEmpty() && !containsAnySkipWord(cellLValue, skipWords)) {
//                    // Store unique values in column M
//                    uniqueValues.add(cellLValue);
//                }
//            }
//        }
//
//        // Store unique values in column N
//        int rowIndex = 0;
//        for (String value : uniqueValues) {
//            Row row = sheet.getRow(rowIndex);
//            if (row == null) {
//                row = sheet.createRow(rowIndex);
//            }
//            Cell cellN = row.createCell(13, CellType.STRING);
//            cellN.setCellValue(value);
//            rowIndex++;
//        }
//    }
    
    private static void extractUniqueValuesFromKtoM(Sheet sheet) {
        DataFormatter dataFormatter = new DataFormatter();
        Set<String> uniqueValues = new LinkedHashSet<>();
        String[] skipWords = {"Name", "Date", "Case", "Type", "Total", "No"};

        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            if (row != null) {
                Cell cellK = row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                String cellKValue = dataFormatter.formatCellValue(cellK);

                if (cellKValue != null && !cellKValue.trim().isEmpty() && !containsAnySkipWord(cellKValue, skipWords)) {
                    // Store unique values in column M
                    uniqueValues.add(cellKValue);
                }
            }
        }

        // Store unique values in column N
        int rowIndex = 0;
        for (String value : uniqueValues) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            Cell cellN = row.createCell(13, CellType.STRING); // Assuming column N corresponds to index 13
            cellN.setCellValue(value);
            rowIndex++;
        }
    }

    // Helper method to check if the value contains any skip word
    private static boolean containsAnySkipWord(String value, String[] skipWords) {
        for (String skipWord : skipWords) {
            if (value.toLowerCase().contains(skipWord.toLowerCase())) {
                return true;
            }
        }
        return false;
    }
}
//5//