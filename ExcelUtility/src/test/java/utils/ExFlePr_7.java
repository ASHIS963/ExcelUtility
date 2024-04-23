package utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExFlePr_7 {
    public static void main(String[] args) {
        try {
            String inputFilePath = "C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_6.xlsx";
            String sheetName = "Sheet1";

            FileInputStream fileInputStream = new FileInputStream(inputFilePath);
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheet(sheetName);

            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }

                Cell cellG = row.getCell(6); // Column G
                Cell cellH = row.getCell(7); // Column H
                Cell cellI = row.getCell(8); // Column I

                if (cellG != null && cellH != null && cellI != null) {
                    String cellGValue = getCellValueAsString(cellG);
                    String cellHValue = getCellValueAsString(cellH);
                    String cellIValue = getCellValueAsString(cellI);

                    // Concatenate values from columns G, H, and I
                    String concatenatedValue = cellGValue + cellHValue + cellIValue;
                    double numericValue = 0;

                    try {
                        numericValue = Double.parseDouble(concatenatedValue);
                    } catch (NumberFormatException e) {
                        // Handle cases where the concatenated value cannot be converted to a number
                    }

                    Cell cellM = row.getCell(12); // Column M
                    if (cellM == null) {
                        cellM = row.createCell(12, CellType.NUMERIC);
                    }
                    cellM.setCellValue(numericValue);
                }
            }

            fileInputStream.close();

            // Save the modified Excel file
            FileOutputStream fileOutputStream = new FileOutputStream(inputFilePath);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

            System.out.println("Excel file has been updated and saved.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else {
            return "";
        }
    }
}
//7//