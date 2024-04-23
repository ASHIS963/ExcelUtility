// writeUniqueDataToColumnP.java

package utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashSet;
import java.util.Set;

public class writeUniqueDataToColumnP {

	public static void main(String[] args) {
        // Specify the path to the Excel file
        String excelFilePath = "C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\tttyty.xlsx";
        // Specify the sheet name
        String sheetName = "Sheet1";
        // Specify the source and target column indices
        int columnIndexSource = 14; // Column O
        int columnIndexTarget = 15; // Column P

        // Call the writeUniqueDataToColumnP method
        writeUniqueDataToColumnP(excelFilePath, sheetName, columnIndexSource, columnIndexTarget);
    }

    public static void writeUniqueDataToColumnP(String excelFilePath, String sheetName, int columnIndexSource, int columnIndexTarget) {
        try {
            // Load the Excel file
            try (FileInputStream fileInputStream = new FileInputStream(excelFilePath)) {
                Workbook workbook = new XSSFWorkbook(fileInputStream);

                // Get the specified sheet
                Sheet sheet = workbook.getSheet(sheetName);

                // Create a set to store unique values in Column O with order preservation
                Set<String> uniqueValues = new LinkedHashSet<>();

                // Iterate through every cell in Column O and add unique values to the set
                int lastRowIndex = findLastNonBlankRow(sheet, columnIndexSource);
                for (int rowIndex = 0; rowIndex <= lastRowIndex; rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    // Get the cell in the source column (Column O)
                    Cell cellSource = row.getCell(columnIndexSource);

                    // Check if the cell is empty in the source column
                    if (isCellEmpty(cellSource)) {
                        continue; // Continue to the next row
                    }

                    // Add the value to the set (ignoring duplicates)
                    uniqueValues.add(cellSource.getStringCellValue().trim());
                }

                // Find the last row in Column P
                int lastRowP = findLastNonBlankRow(sheet, columnIndexTarget);

                // Iterate through the set of unique values and write them to Column P
                for (String uniqueValue : uniqueValues) {
                    // Increment the last row index
                    lastRowP++;

                    // Create a new row if necessary
                    Row rowP = sheet.getRow(lastRowP);
                    if (rowP == null) {
                        rowP = sheet.createRow(lastRowP);
                    }

                    // Write the unique value to the cell in Column P
                    Cell cellTarget = createAndGetCell(rowP, columnIndexTarget);
                    cellTarget.setCellValue(uniqueValue);
                }

                // Save the modified workbook
                try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
                    workbook.write(fileOutputStream);
                }

                // Close the workbook
                workbook.close();

                System.out.println("Unique data from Column O written to Column P in Sheet1.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean isCellEmpty(Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK;
    }

    private static Cell createAndGetCell(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        return cell != null ? cell : row.createCell(columnIndex);
    }

    private static int findLastNonBlankRow(Sheet sheet, int columnIndex) {
        int lastRow = sheet.getLastRowNum();
        for (int i = lastRow; i >= 0; i--) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(columnIndex);
            if (!isCellEmpty(cell)) {
                return i;
            }
        }
        return -1; // Return -1 if no non-blank cell is found in the specified column
    }
}
//vip//