package utils;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class RcntDc_10 {
    // Constants representing column indices in the Excel sheets
    private static final int COLUMN_P_INDEX = 15;
    private static final int COLUMN_L_INDEX = 11;
    private static final int COLUMN_J_INDEX = 9;
    private static final int COLUMN_C_INDEX = 2;
    private static final int COLUMN_M_INDEX = 12;

    public static void main(String[] args) {
        // Specify the file path
        String filePath = "C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_7A.xlsx";

        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Get the sheets from the workbook
            Sheet sheet1 = workbook.getSheet("Sheet1");
            Sheet sheet2 = workbook.getSheet("Sheet2");

            // Process the sheets
            processSheet(sheet1, sheet2);

            // Write the changes back to the file
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }

            System.out.println("Processing completed successfully.");

        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Error during processing: " + e.getMessage());
        }
    }

    // Process the first sheet
    private static void processSheet(Sheet sheet1, Sheet sheet2) {
        // Iterate through rows in Sheet1
        for (int pRowIdx = 0; pRowIdx <= sheet1.getLastRowNum(); pRowIdx++) {
            Row pRow = sheet1.getRow(pRowIdx);
            if (pRow != null) {
                // Get the value of the cell in column P
                Cell cellP = pRow.getCell(COLUMN_P_INDEX);

                // Skip the row if the cell should be skipped
                if (shouldSkipCell(cellP)) {
                    continue; // Continue processing next row in Sheet1
                }

                // Iterate through rows in Sheet1 again
                for (int lRowIdx = 0; lRowIdx <= sheet1.getLastRowNum(); lRowIdx++) {
                    Row lRow = sheet1.getRow(lRowIdx);
                    if (lRow != null) {
                        // Get the value of the cell in column L
                        Cell cellL = lRow.getCell(COLUMN_L_INDEX);

                        // Skip the row if the cell should be skipped
                        if (shouldSkipCell(cellL)) {
                            continue; // Continue processing next row in Sheet1
                        }

                        // Check if values in column P are the same in both sheets
                        if (cellP != null && cellL != null && cellP.toString().equals(cellL.toString())) {
                            // Get the value of the cell in column J
                            Cell cellJ = lRow.getCell(COLUMN_J_INDEX);

                            // Skip the row if column J is "REPEAT"
                            if (cellJ != null && "REPEAT".equalsIgnoreCase(cellJ.toString())) {
                                break; // Continue processing next row in Sheet1
                            }

                            // Get the value of the cell in column C
                            Cell cellC = lRow.getCell(COLUMN_C_INDEX);

                            // Check conditions for processing
                            if (cellC == null || cellC.getCellType() == CellType.BLANK ||
                                    (cellJ == null || cellJ.getCellType() == CellType.BLANK)) {
                                // Find date in the upper direction
                                LocalDate dateSheet1 = findDateInUpperDirection(lRow, COLUMN_C_INDEX);
                                processSheet2(sheet2, cellP, cellJ, dateSheet1, lRow);
                            } else {
                                // Parse date from column C in Sheet1
                                LocalDate dateSheet1 = parseDate(cellC.toString());
                                // Avoid unnecessary call to findDateInUpperDirection if cellC was previously found not null
                                if (dateSheet1 == null) {
                                    // Find date in the upper direction
                                    dateSheet1 = findDateInUpperDirection(lRow, COLUMN_C_INDEX);
                                }
                                processSheet2(sheet2, cellP, cellJ, dateSheet1, lRow);
                            }
                        }
                    }
                }
            }
        }
    }
    private static void processSheet2(Sheet sheet2, Cell cellP, Cell cellJSheet1, LocalDate dateSheet1, Row lRowSheet1) {
        // Iterate through rows in Sheet2
        for (int lRowIdxSheet2 = 0; lRowIdxSheet2 <= sheet2.getLastRowNum(); lRowIdxSheet2++) {
            Row lRowSheet2 = sheet2.getRow(lRowIdxSheet2);
            if (lRowSheet2 != null) {
                // Get the value of the cell in column L
                Cell cellLSheet2 = lRowSheet2.getCell(COLUMN_L_INDEX);

                // Skip the row if the cell should be skipped
                if (shouldSkipCell(cellLSheet2)) {
                    continue; // Continue processing the next row in Sheet2
                }

                // Check if values in column P are the same in both sheets
                if (cellP != null && cellLSheet2 != null &&
                        cellP.toString().equals(cellLSheet2.toString())) {
                    // Get the value of the cell in column C
                    Cell cellCSheet2 = lRowSheet2.getCell(COLUMN_C_INDEX);

                    // Check conditions for processing
                    if (cellCSheet2 == null || cellCSheet2.getCellType() == CellType.BLANK) {
                        // Find date in the upper direction for Sheet2
                        LocalDate dateSheet2 = findDateInUpperDirection(lRowSheet2, COLUMN_C_INDEX);

                        // Compare dates from Sheet1 and Sheet2
                        if (dateSheet1 != null && dateSheet2 != null) {
                            if (dateSheet2.isAfter(dateSheet1)) {
                                updateSheet2CellValues(lRowSheet2, cellP);
                            }
                            // Update Sheet1 if the date in Sheet1 is more recent
                            if (dateSheet1.isAfter(dateSheet2)) {
                                updateSheet1CellValues(lRowSheet1, cellP);
                            }
                        }
                    } else {
                        // Parse date from column C in Sheet2
                        LocalDate dateSheet2 = parseDate(cellCSheet2.toString());
                        // Avoid unnecessary call to findDateInUpperDirection if cellCSheet2 was previously found not null
                        if (dateSheet2 == null) {
                            // Find date in the upper direction for Sheet2
                            dateSheet2 = findDateInUpperDirection(lRowSheet2, COLUMN_C_INDEX);
                        }

                        // Compare dates from Sheet1 and Sheet2
                        if (dateSheet1 != null && dateSheet2 != null) {
                            if (dateSheet2.isAfter(dateSheet1)) {
                                updateSheet2CellValues(lRowSheet2, cellP);
                            }
                            // Update Sheet1 if the date in Sheet1 is more recent
                            if (dateSheet1.isAfter(dateSheet2)) {
                                updateSheet1CellValues(lRowSheet1, cellP);
                            }
                        }
                    }
                }
            }
        }
    }


    // Update cells in Sheet1 based on conditions
    private static void updateSheet1CellValues(Row lRow, Cell cellP) {
        // Update cells in Sheet1
        Cell cellGSheet1 = lRow.getCell(6); // Column G
        Cell cellHSheet1 = lRow.getCell(7); // Column H
        Cell cellISheet1 = lRow.getCell(8); // Column I
        Cell cellJSheet1 = lRow.getCell(COLUMN_J_INDEX);

        if ((cellGSheet1 != null && cellGSheet1.getCellType() != CellType.BLANK) ||
                (cellHSheet1 != null && cellHSheet1.getCellType() != CellType.BLANK) ||
                (cellISheet1 != null && cellISheet1.getCellType() != CellType.BLANK)) {
            if (cellGSheet1 != null) {
                cellGSheet1.setCellValue("");
            }
            if (cellHSheet1 != null) {
                cellHSheet1.setCellValue(0);
                if (cellJSheet1 != null) {
                    cellJSheet1.setCellValue("REPEAT");
                }
            }
            if (cellISheet1 != null) {
                if (cellJSheet1 == null) {
                    cellJSheet1.setCellValue("REPEAT");
                }
            }
            Cell cellMSheet1 = lRow.getCell(COLUMN_M_INDEX);
            if (cellMSheet1 != null) {
                cellMSheet1.setCellValue("0");
                CellStyle redCellStyle = lRow.getSheet().getWorkbook().createCellStyle();
                redCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                redCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellMSheet1.setCellStyle(redCellStyle);
                System.out.println("Sheet1" + cellP);
            }
        }
    }

    // Update cells in Sheet2 based on conditions
    private static void updateSheet2CellValues(Row lRowSheet2, Cell cellP) {
        // Update cells in Sheet2
        Cell cellGSheet2 = lRowSheet2.getCell(6); // Column G in Sheet2
        Cell cellHSheet2 = lRowSheet2.getCell(7); // Column H in Sheet2
        Cell cellISheet2 = lRowSheet2.getCell(8); // Column I in Sheet2
        Cell cellJSheet2 = lRowSheet2.getCell(COLUMN_J_INDEX);

        if ((cellGSheet2 != null && cellGSheet2.getCellType() != CellType.BLANK) ||
                (cellHSheet2 != null && cellHSheet2.getCellType() != CellType.BLANK) ||
                (cellISheet2 != null && cellISheet2.getCellType() != CellType.BLANK)) {
            if (cellGSheet2 != null) {
                cellGSheet2.setCellValue("");
            }
            if (cellHSheet2 != null) {
                cellHSheet2.setCellValue(0);
            }
            if (cellISheet2 != null) {
                cellISheet2.setCellValue("");
            }
            Cell cellMSheet2 = lRowSheet2.createCell(COLUMN_M_INDEX);
            cellMSheet2.setCellValue("0");
            CellStyle redCellStyle = lRowSheet2.getSheet().getWorkbook().createCellStyle();
            redCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            redCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            lRowSheet2.getCell(COLUMN_L_INDEX).setCellStyle(redCellStyle);
            System.out.println("Sheet1" + cellP);
            if (cellJSheet2 != null) {
                cellJSheet2.setCellValue("REPEAT");
                System.out.println("Sheet2" + cellP);
            }
        }
    }

    // Check if a cell should be skipped
    private static boolean shouldSkipCell(Cell cell) {
        // Check if the cell should be skipped
        return cell == null || cell.getCellType() == CellType.BLANK ||
                (cell.getCellType() == CellType.STRING && isSkipWord(cell.toString()));
    }

    // Check if a word should be skipped
    private static boolean isSkipWord(String word) {
        // Check if the word should be skipped
        String[] skipWords = {"Name", "Date", "Case", "Type", "Total", "No"};
        for (String skipWord : skipWords) {
            if (word.equalsIgnoreCase(skipWord)) {
                return true;
            }
        }
        return false;
    }

    // Parse the date string with different formats
    private static LocalDate parseDate(String dateStr) {
        // Parse the date string
        String[] possibleFormats = {"d.M.yy", "dd.MM.yy", "dd.MM.yyyy", "d.MM.yy",
                "dd.MM.yy", "d.MM.yyyy", "dd.MM.yyyy"};

        for (String format : possibleFormats) {
            try {
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern(format);
                return LocalDate.parse(dateStr, formatter);
            } catch (Exception e) {
                // Try the next format
            }
        }

        return null; // Handle parsing failure appropriately
    }

    // Find the date in the upper direction of the current cell
    private static LocalDate findDateInUpperDirection(Row row, int columnIndex) {
        // Get the current row index
        int rowIndex = row.getRowNum();

        // Iterate in the upper direction starting from the current row
        for (int i = rowIndex; i >= 0; i--) {
            Row currentRow = row.getSheet().getRow(i);

            // Check if the current row is not null
            if (currentRow != null) {
                Cell currentCell = currentRow.getCell(columnIndex);

                // Check if the current cell is not blank, empty, or null
                if (currentCell != null && currentCell.getCellType() != CellType.BLANK &&
                        !(currentCell.getCellType() == CellType.STRING && currentCell.getStringCellValue().trim().isEmpty())) {
                    // Parse the date from the current cell
                    return parseDate(currentCell.toString());
                }
            }
        }
        return null; // Return null if no date is found in the upper direction
    }
}
// final //10  // update in both sheet.