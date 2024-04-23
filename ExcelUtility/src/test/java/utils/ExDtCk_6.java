package utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;

public class ExDtCk_6 {
    public static void main(String[] args) {
        try {
            // Specify the path to the Excel file
            String excelFilePath = "C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_5.xlsx";

            // Load the Excel file
            FileInputStream fileInputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheet("Sheet1");

            // Define skip words
            String[] skipWords = {"Name", "Date", "Case", "Type", "Total", "No"};

            // Create a data formatter to handle cell values
            DataFormatter dataFormatter = new DataFormatter();

            // Iterate through the rows of the sheet
            for (int rowIndexN = 0; rowIndexN < sheet.getPhysicalNumberOfRows(); rowIndexN++) {
                Row rowN = sheet.getRow(rowIndexN);
                if (rowN == null) {
                    continue; // Skip empty rows
                }

                // Get the value in column N for the current row
                Cell cellN = rowN.getCell(13); // Column N, 0-based index
                if (cellN == null) {
                    continue; // Skip rows with empty cells
                }

                String valueN = dataFormatter.formatCellValue(cellN);

                // Map to store the count of occurrences for each value in column L
                Map<String, Integer> valueCountMap = new HashMap<>();

                // Iterate over column L from 0th row to the last row with data
                for (int rowIndexL = 0; rowIndexL < sheet.getPhysicalNumberOfRows(); rowIndexL++) {
                    Row rowL = sheet.getRow(rowIndexL);
                    if (rowL == null) {
                        continue; // Skip empty rows
                    }

                    // Get the value in column L for the current row
                    Cell cellL = rowL.getCell(11); // Column L, 0-based index
                    if (cellL == null) {
                        continue; // Skip rows with empty cells
                    }

                    String valueL = dataFormatter.formatCellValue(cellL);

                    // Check if the valueL contains any skip words
                    boolean isSkipWord = false;
                    for (String skipWord : skipWords) {
                        if (valueL.contains(skipWord)) {
                            isSkipWord = true;
                            break; // No need to check further
                        }
                    }

                    if (isSkipWord) {
                        continue; // Skip rows with skip words in column L
                    }

                    // Increment the count for the value in column L
                    int count = valueCountMap.getOrDefault(valueL, 0);
                    valueCountMap.put(valueL, count + 1);

                    // Compare valueN with valueL and perform tasks for the second and subsequent occurrences
                    if (valueN.equalsIgnoreCase(valueL) && valueCountMap.get(valueL) >= 2) {
                        // Perform tasks for the 2nd occurrence onwards
                        Cell cellJ = rowL.getCell(9);  // Column J
                        Cell cellG = rowL.getCell(6);  // Column G
                        Cell cellH = rowL.getCell(7);  // Column H

                        // Check the J cell and update G, H, and J as needed
                        if (cellJ == null || (!dataFormatter.formatCellValue(cellJ).equalsIgnoreCase("REPEAT"))) {
                            if (cellG == null) {
                                cellG = rowN.createCell(6, CellType.BLANK);
                            }
                            cellG.setCellValue("");
                            if (cellH == null) {
                                cellH = rowN.createCell(7, CellType.NUMERIC);
                            }
                            cellH.setCellValue(0);
                            if (cellJ == null) {
                                cellJ = rowN.createCell(9, CellType.STRING);
                            }
                            cellJ.setCellValue("REPEAT");
                            System.out.println("Match found: " + valueL + " = " + valueN);
                        }
                    }
                }
            }

            // Save the updated Excel file as "tttyty.xlsx" in the same directory
            String outputFilePath = "C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_6.xlsx";
            FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

            // Close the input stream
            fileInputStream.close();
            System.out.println("Excel file has been updated and saved as OutPut_6.xlsx.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
//6//