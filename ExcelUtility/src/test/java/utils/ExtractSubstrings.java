package utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExtractSubstrings {
    public static void main(String[] args) {
    	
        try {
        	ZipSecureFile.setMinInflateRatio(0.001);
            FileInputStream file = new FileInputStream("C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_7.xlsx");
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0); // Assuming first sheet

            // Define regex patterns for valid file prefixes
            String[] prefixFormats = {"ABLAPL","ADMLS","AHO","ARBA(ICA)","ARBA","ArbitrationAppeal","ARBP(ICA)","ARBP","AS","BKGP","BLAPL","CEREF","CMAPL","CMP","CMPA","CMPAT","CMPMC","CO","COA","COAPL","COCAS","CONTAC","CONTAR","CONTC(CP)","CONTC(CPB)","CONTC(CPC)","CONTC(CPS)","CONTC","CONTR","COPET","CR","CRA","CRLA","CRLLP","CRLMA","CRLMC","CRLMP","CRLREF","CRLREV","CRLTR","CRMC","CRP","CRREF","CRREV","CS(OS)","CS","CUSREF","CVA","CVREF","CVREV","CVRVW","DREF","DSREF","EC","EDA","EDR","ELA","ELPET","EP","EXFA","EXOS","EXP","EXSA","FA(OS)","FA","FAO","GA","GCRLA","GTA","GTR","GUAP","IA","INSA","INSREF","INTEST","IP(M)","IPAPL","ITA","ITR","JCRA","JCRLA","JCRLMC","JCRLRV","JCRMC","JCRREV","LAA","LAREF","LPA","MA","MAC","MACA","MATA","MATCAS","MATREF","MFA","MJC","MREF","MSA","MSREF","NM","OCRMC","OJC","OREF","OS","OTAPL","OTC","OTR","OVTA","RCC","RCFA","RCREV","RCSA","RFA","RMC","RPFAM","RSA","RVWPET(RP)","RVWPET(RPB)","RVWPET(RPC)","RVWPET(RPS)","RVWPET","SA","SAO","SCA","SCLP","SJC","SM","SPA","SPJC","STAPL","STREF","STREV","TA","TEST","TMC","TREV","TRP(C)","TRPCRL","WA","WP(C)","WPC(OA)","WPC(OAB)","WPC(OAC)","WPC(OAP)","WPC(OAPB)","WPC(OAPC)","WPC(OAPP)","WPC(OAPPB)","WPC(OAPPC)","WPC(OAPPS)","WPC(OAPS)","WPC(OAS)","WPC(T)","WPC(TA)","WPC(TAB)","WPC(TAC)","WPC(TAS)","WPCRL","WTA","WTR",/* Add more formats as needed */};

            
            String[] specialChar = {"!", "@", "#", "$", "%", "^", "&", "*", "-", "+", "=", "[", "]", "{", "}", "|", ";", ":", "`", ",", ".", "/", "<", ">", "?", "~", "'", "\\","'"};


            // Process each row in the sheet
         // Create a cell style for center alignment
            CellStyle centerStyle = workbook.createCellStyle();
            centerStyle.setAlignment(HorizontalAlignment.CENTER); // Set alignment
            centerStyle.setVerticalAlignment(VerticalAlignment.CENTER); // Set vertical alignment

            // Apply the style to all cells in the sheet
            for (Row row : sheet) {
                Cell cell = row.getCell(13); // Assuming column "N" is the 14th column (0-indexed)
                if (cell == null || cell.getCellType() != CellType.STRING) {
                    continue; // Skip if cell is empty or not a string
                }
                String value = cell.getStringCellValue();

                // Process the value according to the provided criteria
                String processedValue = processValue(value, prefixFormats, specialChar);

                // Set the processed value in column "O" of the same row
                int columnIndex = 14; // Assuming column "O" is the 15th column (0-indexed)
                Cell newCell = row.createCell(columnIndex);
                newCell.setCellValue(processedValue);

                // Apply the pre-defined cell style for center alignment
                newCell.setCellStyle(centerStyle);
            }

            
            // Write to output Excel file
            FileOutputStream outFile = new FileOutputStream("C:\\Users\\anthem\\eclipse-workspace\\ExcelUtility\\data\\OutPut_7A.xlsx");
            workbook.write(outFile);
            outFile.close(); // Close the file output stream
            workbook.close();
            System.out.println("Extraction completed successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    private static String processValue(String value, String[] prefixFormats,String[]specialChar) {
        StringBuilder processedValue = new StringBuilder();

        // Variables for tracking parts of the value
        StringBuilder filePrefix = new StringBuilder();
        StringBuilder fileNo = new StringBuilder();
        StringBuilder fileYear = new StringBuilder();
        StringBuilder residualBody = new StringBuilder();

        boolean yearFound = false; // Flag to indicate if year digits are being found
        int partchk=1;
        int partCounter = 0;
        for (char ch : value.toCharArray()) {
            if (ch == '_') {
                partCounter++;
                partchk=0;
                continue;
            }

            switch (partCounter) {
                case 0:
                    filePrefix.append(ch);
                    break;
                case 1:
                    fileNo.append(ch);
                    break;
                case 2:
                    if (Character.isDigit(ch)) {
                        if (!yearFound) {
                            fileYear.append(ch);
                        } else {
                            // Stop iteration if another digit found after year is complete
                            partCounter++;
                        }
                    } else {
                        // Stop iteration if encounter a non-digit character after year
                        partCounter++;
                    }
                    break;
                default:
                    residualBody.append(ch);
                    break;
            }
//			if((partCounter == 0)&&(partchk==1))
//			{System.out.println("check File name: "+value);
//			partchk--;}
            
            // Check if the current character indicates the end of the year digits
            if (fileYear.length() == 4 || (fileYear.length() == 2 && !Character.isDigit(ch))) {
                yearFound = true;
            }
        }

        // Processing File_Prefix
        String prefix = filePrefix.toString();
        StringBuilder alphaPrefix = new StringBuilder(); // StringBuilder to store alphabetical characters
        for (char ch : prefix.toCharArray()) {
            if (Character.isLetter(ch)) {
            	alphaPrefix.append(Character.toUpperCase(ch)); // Append uppercase version // Append alphabetical characters to alphaPrefix
            }
        }

        // Match file prefix with format prefixes
        boolean filePrefixMatched = false;
        for (String format : prefixFormats) {
            // Extract alphabetical characters from the format prefix
            StringBuilder alphaFormatPrefix = new StringBuilder();
            for (char ch : format.toCharArray()) {
                if (Character.isLetter(ch)) {
                    alphaFormatPrefix.append(ch); // Append alphabetical characters to alphaFormatPrefix
                }
            }

            // Check if file prefix matches the format prefix
            if (alphaFormatPrefix.toString().equals(alphaPrefix.toString())) {
                // If match found, append the format to processed value
                processedValue.append(format);
                filePrefixMatched = true;
                break;
            }
        }

        // If none of the formats matched the file prefix, print the file prefix
        if (!filePrefixMatched) {
            System.out.println("Check file prefix	:" + value);
        }

        // Processing File_Year
        String year = fileYear.toString();
        if (year.length() == 2) {
            int yearValue = Integer.parseInt(year);
           year = (yearValue >= 1 && yearValue <= 24) ? "20" + year : "19" + year;
//            if(yearValue >= 1 && yearValue <= 24) {
//            	year ="20" + year;
//            }else {System.out.println("issue in year		:"+value);}
        }
        //BEFORE RESIDUAL COUNT FOR LCR
        int conchek = 1;
        for (int i = 0; i < residualBody.length() - 1; i++) {
            char currentChar = residualBody.charAt(i);

            // Check if the current character is a digit
            if (Character.isDigit(currentChar)) {
                char nextChar = residualBody.charAt(i + 1);
                // Check if the next character is also a digit
                if (Character.isDigit(nextChar)) {
                    // Extract and process the two-digit number
                    String twoDigitNumber = residualBody.substring(i, i + 2);
                    // Process the two-digit number as needed
                    if(conchek>0) {System.out.println("issue number/year	:" + value);
                    conchek--;}
                }
            }
        }
        
        // Processing ResidualBody
        residualBody.setLength(0); // Clear the residualBody before appending new values
        if (value.toUpperCase().contains("LCR")) {
            residualBody.append("LCR");
        } else if (value.toUpperCase().contains("TCR")) {
            residualBody.append("TCR");
        } else if (value.toUpperCase().contains("P.B")) {
            residualBody.append("PB");
        } else if (value.toUpperCase().contains("PAPER BOOK")) {
            residualBody.append("PB");
        } else if (value.toUpperCase().contains("PAPERBOOK")) {
            residualBody.append("PB");
        } else if (value.toUpperCase().contains("PAPER")) {
            residualBody.append("PB");
        } else if (value.toUpperCase().contains("BOOK")) {
            residualBody.append("PB");
        } else if (value.toUpperCase().contains("PB")) {
            residualBody.append("PB");
        }

        // Append parts to form the processed value       
//        processedValue.append("_").append(fileNo).append("_").append(year);      
//        if (residualBody.length() > 0) {
//           processedValue.append("_").append(residualBody);
//       }
  
     // Append parts to form the processed value
        String fileNoString = fileNo.toString().trim();
        String yearString = year.toString().trim();

        if (fileNoString.isEmpty() || yearString.isEmpty()) {
            System.out.println("Problem in String	:" + value);
            return ""; // Return an empty string to signify the problem
        }

        processedValue.append("_").append(fileNoString).append("_").append(yearString);

        if (residualBody.length() > 0) {
            processedValue.append("_").append(residualBody);
        }


         
         
        boolean containsSpecialChar = false;
        for (char ch : processedValue.toString().toCharArray()) {
            for (String sch : specialChar) {
                if (ch == sch.charAt(0)) {
                    containsSpecialChar = true;
                    break;
                }
            }
            if (containsSpecialChar) {
                System.out.println("Problem in String	:" + processedValue);
                break; // No need to continue checking once a special character is found
            }
        }

        return processedValue.toString();
    }}
/* 
 * before process replace"__"with"_" .
 * replace"/"with"_" and remove blank space.
 * Check file prefix	;here filePrefix problem.
 * issue number/year	:here file number or year problem.
 * Problem in String	:here may be a special character or a null 
 */
//after (7)but its not 8 (A)