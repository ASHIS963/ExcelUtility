package utils;

import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

public class ExUp_11 {
    public static void main(String[] args) throws IOException {
        String excelPath = "./data/OutPut_7A.xlsx"; // Update with the path of the modified Excel file
        long cellcount = 0; // Initialize cellcount to 0
        long sumvalue = 0; // Initialize sumvalue to 0
        Cell jCell = null; // Declare jCell outside the "total" row handling block
        try {
            ExcelUtils excel = new ExcelUtils(excelPath, "Sheet1");
            Sheet sheet = excel.getSheet();

            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);

                // Add a null check for the row object
                if (row == null) {
                    continue; // Skip this row if it's null
                }

                Cell kCell = row.getCell(10);

                if (kCell == null || kCell.getCellType() == CellType.BLANK) {
                    // Skip the row if K cell is blank, empty, or null
                    continue;
                }

                String kCellValue = excel.getCellStringValue(kCell).toLowerCase();

                if (kCellValue.contains("total")) {
                    // Handle "total" row
                    Cell fCell = row.getCell(5); // Column F cell
                    Cell gCell = row.getCell(6); // Column G cell
                    Cell hCell = row.getCell(7); // Column H cell
                    Cell iCell = row.getCell(8); // Column I cell
                    jCell = row.getCell(9); // Column J cell
                    Cell QCell = row.getCell(16); // Column N cell
                    Cell RCell = row.getCell(17); // Column Q cell

                    if (fCell != null) {
                        // Clear the previous value in column F
                        fCell.setCellValue("");
                        // Set cellcount in column F
                        fCell.setCellValue(cellcount);

                    }

                    if (gCell != null) {
                        // Clear the previous value in column G
                        gCell.setCellValue("");
                    }

                    if (hCell != null) {
                        // Clear the previous value in column H
                        hCell.setCellValue("");
                    }

                    if (iCell != null) {
                        // Clear the previous value in column I
                        iCell.setCellValue("");
                    }

                    if (jCell != null) {
                        // Clear the previous value in column J
                        jCell.setCellValue("");
                        // Set sumvalue in column J
                        jCell.setCellValue(sumvalue);
                        if (RCell == null) {
                            // Create column O cell if it doesn't exist
                            RCell = row.createCell(17, CellType.NUMERIC);
                        }

                        // Set sumvalue in column Q
                        RCell.setCellValue(sumvalue);
                    }

                    if (QCell == null) {
                        // Create column N cell if it doesn't exist
                        QCell = row.createCell(16, CellType.NUMERIC);
                    }
                    // Set sumvalue in column N
                    QCell.setCellValue(cellcount);
                    // Reset cellcount and sumvalue
                    System.out.println("file: "+cellcount+" & "+"count: "+sumvalue);
                    cellcount = 0;
                    sumvalue = 0;
                } else {
                    // Check if the cell value in column M is numeric
                    Cell mCell = row.getCell(12); // Column M cell

                    if (mCell != null && mCell.getCellType() == CellType.NUMERIC) {
                        double mValue = mCell.getNumericCellValue();

                        if (mValue != 0) {
                            // Add the value to sumvalue and increment cellcount
                            sumvalue += (long) mValue;
                            cellcount++;
                            //System.out.println(cellcount);
                        }
                        
                    }
                } 
            }
            
            // Save the changes to the Excel file
            String outputPath = "./data/Final_Excel.xlsx"; // Update the output file path
            excel.saveExcel(outputPath);

            System.out.println("Processed values in column M as numbers, updated columns F, G, H, I, J, Q, and R for 'total' rows, and performed tax calculation.");
            System.out.println("Sheet1 update successfuly.");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Reverted changes due to an error.");
        }
    }
}
//11// update the single sheet i.e: Sheet1.