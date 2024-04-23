// OtoPdatatrans.java

package utils;

public class OtoPdat_9 {
    public static void main(String[] args) {
        // Specify the path to the Excel file
        String excelFilePath = "C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_7A.xlsx";
        // Specify the sheet name
        String sheetName = "Sheet1";
        // Specify the source and target column indices
        int columnIndexSource = 14; // Column O
        int columnIndexTarget = 15; // Column P

        // Call the writeUniqueDataToColumnP method
        writeUniqueDataToColumnP.writeUniqueDataToColumnP(excelFilePath, sheetName, columnIndexSource, columnIndexTarget);
    }
}
//9// run it after compare all sheets with it COMPARE ONCE.