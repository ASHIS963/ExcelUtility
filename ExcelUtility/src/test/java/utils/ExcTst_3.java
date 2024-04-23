package utils;

import java.io.IOException;

public class ExcTst_3 {
    public static void main(String[] args) {
        String excelPath = "./data/rtr.xlsx"; // Update with your Excel file path
        String outputPath = "./data/OutPut_3.xlsx"; // Update with the output file path

        try (ExcelUtils excel = new ExcelUtils(excelPath, "Sheet1")) {
            // Call the concatenateEFGAndStoreInColumnK method
            excel.concatenateEFGAndStoreInColumnK();
            System.out.println("poorthfyfhj   ");
            // Call the processColumnK method to process column K
            excel.processColumnK();
            System.out.println("Error during K.");
            // Call the processColumnG method to process column G
            excel.processColumnG();
           
            // Save the changes to the Excel file
            excel.saveExcel(outputPath);

            System.out.println("Concatenated, processed values in column K, and processed values in column G.");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Error during processing.");
        }
    }
}
//3//