package utils;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EelUtiTc_4 {
    public static void main(String[] args) throws IOException {
        String excelPath = "./data/OutPut_3.xlsx"; // Update with the path of the modified Excel file

        try {
            ExcelUtils excel = new ExcelUtils(excelPath, "Sheet1");

            Sheet sheet = excel.getSheet();

            long cellcount = 0; // Initialize cellcount to 0
            long sumvalue = 0; // Initialize sumvalue to 0

            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                
                // Add a null check for the row object
                if (row == null) {
                    continue; // Skip this row if it's null
                }

                Cell kCell = row.getCell(10);
                Cell mCell = row.getCell(12); // Column M cell

                if (kCell == null || kCell.getCellType() == CellType.BLANK) {
                    // Skip the row if K cell is blank, empty, or null
                    continue;
                }
                String kCellValue = excel.getCellStringValue(kCell).toLowerCase();
                if (!excel.shouldSkip(kCellValue)) {
                    // Concatenate the cells of columns G, H, and I and store in column M
                    String concatenatedValue = "";
                    Cell gCell = row.getCell(6);
                    Cell hCell = row.getCell(7);
                    Cell iCell = row.getCell(8);

                    if (gCell != null) {
                        concatenatedValue += excel.getCellStringValue(gCell);
                    }

                    if (hCell != null) {
                        concatenatedValue += excel.getCellStringValue(hCell);
                    }

                    if (iCell != null) {
                        concatenatedValue += excel.getCellStringValue(iCell);
                    }

                    if (mCell == null) {
                        mCell = row.createCell(12); // Create a new column M cell if it doesn't exist
                    }

                    mCell.setCellValue(concatenatedValue); // Set the concatenated value in column M

                    
                    }
                }
            

            // Additional loop to check column M values           
            // Save the changes to the Excel file
            String outputPath = "./data/OutPut_4.xlsx"; // Update the output file path
            excel.saveExcel(outputPath);

            System.out.println("Processed values in column M, updated columns F, J, N, and O for 'total' rows, and performed tax calculation.");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Reverted changes due to an error.");
        }
    }
}
//4//