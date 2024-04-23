package utils;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class cSUq_8 {
    public static void main(String[] args) {
        try {
            // Specify the path to the Excel file
            String excelFilePath = "C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_7A.xlsx";

            // Load the Excel file
            try (FileInputStream fileInputStream = new FileInputStream(excelFilePath)) {
                Workbook workbook = new XSSFWorkbook(fileInputStream);

                // Get "Sheet1" and "Sheet2" from the workbook
                Sheet sheet1 = workbook.getSheet("Sheet1");
                Sheet sheet2 = workbook.getSheet("Sheet2");

                // Find the last non-blank row in both sheets for column N
                int lastRow1 = findLastNonBlankRow(sheet1, 13); // Column N
                int lastRow2 = findLastNonBlankRow(sheet2, 13); // Column N

                // Define the column indices for comparison (Column N and O, 0-based indices)
                int columnIndexN = 13;
                int columnIndexO = 14; // Column O

                // Create a CellStyle with red font
                CellStyle redFontStyle = createRedFontStyle(workbook);

                // Track the current column index in Sheet1 for writing to Column O
                int currentColumnIndexO = columnIndexO;

                // Iterate through every non-blank cell in column N of both sheets and compare the data
                for (int rowIndex1 = 0; rowIndex1 <= lastRow1; rowIndex1++) {
                    Row row1 = sheet1.getRow(rowIndex1);
                    // Get the cell in column N for Sheet1
                    Cell cell1N = row1.getCell(columnIndexN);

                    // Check if the cell is empty in Sheet1
                    if (isCellEmpty(cell1N)) {
                        System.out.println("Cell is empty in Sheet1, row " + (rowIndex1 + 1));
                        continue; // Continue to the next row
                    }

                    String value1N = cell1N.getStringCellValue().trim();
                    boolean matchFound = false;

                    for (int rowIndex2 = 0; rowIndex2 <= lastRow2; rowIndex2++) {
                        Row row2 = sheet2.getRow(rowIndex2);

                        // Check if rowIndex2 is within the valid range
                        if (rowIndex2 > lastRow2) {
                            continue; // Continue to the next iteration
                        }

                        // Get the cell in column N for Sheet2
                        Cell cell2N = row2.getCell(columnIndexN);

                        // Check if the cell is empty in Sheet2
                        if (isCellEmpty(cell2N)) {
                            System.out.println("Cell is empty in Sheet2, row " + (rowIndex2 + 1));
                            continue; // Continue to the next row
                        }

                        String value2N = cell2N.getStringCellValue().trim();

                        if (value1N.equals(value2N)) {
                            // Set the font color to red for matching cells in both sheets
                            cell1N.setCellStyle(redFontStyle);

                            // Write the matched value to the cell in column O of Sheet1
                            Cell cell1O = createAndGetCell(row1, columnIndexO);
                            cell1O.setCellValue(value1N);

                            matchFound = true;
                            break; // Exit the loop once a match is found
                        }
                    }

                    if (matchFound) {
                        System.out.println("Match found at Sheet1, row " + (rowIndex1 + 1));
                        System.out.println("Matched Value (Column N): " + value1N);
                    }
                }

                // Save the modified workbook
                try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
                    workbook.write(fileOutputStream);
                }

                // Close the workbook
                workbook.close();

                System.out.println("Comparison completed. Check the Excel file for highlighted cells and matched values in Column O of Sheet1.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static CellStyle createRedFontStyle(Workbook workbook) {
        CellStyle redFontStyle = workbook.createCellStyle();
        Font redFont = workbook.createFont();
        redFont.setColor(IndexedColors.RED.getIndex());
        redFontStyle.setFont(redFont);
        return redFontStyle;
    }

    private static boolean isCellEmpty(Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK;
    }

    private static Cell createAndGetCell(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);

        // Check if the cell is null and create a new cell if necessary
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }

        return cell;
    }

    private static int findLastNonBlankRow(Sheet sheet, int columnIndex) {
        int lastRow = sheet.getLastRowNum();
        for (int i = lastRow; i >= 0; i--) {
            Row row = sheet.getRow(i);

            // Check if the row is null and create a new row if necessary
            if (row == null) {
                row = sheet.createRow(i);
            }

            Cell cell = row.getCell(columnIndex);
            if (!isCellEmpty(cell)) {
                return i;
            }
        }
        return -1; // Return -1 if no non-blank cell is found in the specified column
    }
}

//8// before run set 2 sheet name.(Sheet1 & Sheet2).