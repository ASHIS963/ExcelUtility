package utils;

import java.io.File;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelProcessor {

    public static void main(String[] args) {
        ZipSecureFile.setMinInflateRatio(0);
        String inputFile = "C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_7A.xlsx";
        String outputFile = "C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\FinalExcel.xlsx";

        try {
            processExcel(inputFile, outputFile);
            System.out.println("Processing completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Error occurred while processing the Excel file.");
        }
    }

    private static void processExcel(String inputFile, String outputFile) throws IOException {
        FileInputStream inputStream = new FileInputStream(inputFile);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        for (int nRow = 0; nRow <= sheet.getLastRowNum(); nRow++) {
            XSSFRow nRowObject = sheet.getRow(nRow);
            XSSFCell nCell = nRowObject.getCell(13); // Column N
            XSSFCell oCell = nRowObject.getCell(14); // Column O

            if (nCell != null && oCell != null) {
                String nValue = getCellValueAsString(nCell);
                String oValue = oCell.getStringCellValue();

                for (int lRow = 0; lRow <= sheet.getLastRowNum(); lRow++) {
                    XSSFRow lRowObject = sheet.getRow(lRow);
                    XSSFCell lCell = lRowObject.getCell(11); // Column L

                    if (lCell != null) {
                        String lValue = lCell.getStringCellValue();

                        if (nValue.equals(lValue)) {
                            XSSFCell kCell = lRowObject.getCell(10); // Column K
                            XSSFCell eCell = lRowObject.getCell(4);  // Column E
                            XSSFCell fCell = lRowObject.getCell(5);  // Column F

                            if (kCell != null && (eCell != null || fCell != null)) {
                                String kValue = getCellValueAsString(kCell);
                                String eValue = eCell != null ? getCellValueAsString(eCell) : null;
                                String fValue = fCell != null ? getCellValueAsString(fCell) : null;

                                if (kValue != null && kValue.equals(nValue)) {
                                    if (eValue != null && eValue.equals(nValue)) {
                                        eCell.setCellValue(oValue);
                                    }
                                    if (fValue != null && fValue.equals(nValue)) {
                                        fCell.setCellValue(oValue);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        inputStream.close();

        FileOutputStream outputStream = new FileOutputStream(outputFile);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
        System.out.println("PROCESS COMPLETE");
    }

    private static String getCellValueAsString(XSSFCell cell) {
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            // Handle numeric values
            return String.valueOf(cell.getNumericCellValue());
        }
        return "";
    }
}
// file (B)