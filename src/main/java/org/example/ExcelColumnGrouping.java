package org.example;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelColumnGrouping {
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\5CG6105SVT\\Desktop\\LTE-SM-2023-Allocation-List-Lots-10111219-Nikka-Trading-1-Copy.xlsx";

        try {
            // Load Excel file
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(fis);

            // Assuming there's only one sheet, you may need to iterate over sheets if necessary
            Sheet sheet = workbook.getSheetAt(0);

            // Create maps to store groups for columns C
            Map<String, StringBuilder> groupsC = new HashMap<>();

            // Specify the column indices for columns C (assuming zero-based index)
            int columnIndexC = 2; // Column C is index 2

            // Iterate over rows in columns C, starting from row 4 (index 3)
            for (int rowIndex = 3; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                // Process column C
                Cell cellC = row.getCell(columnIndexC);
                if (cellC != null) {
                    String cellValueC = getCellValue(cellC);
                    groupsC.computeIfAbsent(cellValueC, k -> new StringBuilder()).append(rowIndex).append(",");
                }

            }

            // Output the groups for column C
            System.out.println("Groups for Column C:");
            int count = 0; // Counter variable
            for (Map.Entry<String, StringBuilder> entry : groupsC.entrySet()) {
                System.out.println("Value '" + entry.getKey() + "': Rows " + entry.getValue());
                count++;
                if (count >= 2) {
                    break; // Break out of the loop once the limit is reached
                }
            }


            // Close the workbook and file input stream
            workbook.close();
            fis.close();
        } catch (IOException | EncryptedDocumentException e) {
            e.printStackTrace();
        }
    }

    // Helper method to get the cell value as string
    private static String getCellValue(Cell cell) {
        String cellValue = "";
        // Check the cell type and handle accordingly
        switch (cell.getCellType()) {
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case NUMERIC:
                // Handle numeric values
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            // Add cases for other cell types as needed
        }
        return cellValue;
    }
}
