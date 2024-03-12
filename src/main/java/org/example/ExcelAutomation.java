package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class ExcelAutomation {
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\5CG6105SVT\\Desktop\\LTE-SM-2023-Allocation-List-Lots-10111219-Nikka-Trading-1-Copy.xlsx";

        // Path to your existing Excel template file
        String templateFilePath = "res/files/receiptTemplate.xlsx";

        // Path to save the new Excel file
        String newFilePath = "C:\\Users\\5CG6105SVT\\Desktop\\newFile.xlsx";

        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0); // Assuming it's the first sheet

            // Get the fourth row
            Row row4 = sheet.getRow(3); // Index 3 represents the fourth row

            // Initialize list to store values of columns 0 to 4 in row 4
            List<String> row4Values = new ArrayList<>();

            // Get values of columns 0 to 4 in row 4
            for (int colIndex = 0; colIndex <= 4; colIndex++) {
                Cell cell = row4.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // Ensure cell is not null
                System.out.println("Processing cell in column " + (colIndex + 1) + " of row 4");
                System.out.println("Cell value: " + getCellValueAsString(cell)); // Print cell value for debugging
                row4Values.add(getCellValueAsString(cell)); // Add cell value to list
            }

            int mergedRegionsCount = sheet.getNumMergedRegions(); // Get the total number of merged regions

            // Create a list to store merged regions
            List<CellRangeAddress> mergedRegions = new ArrayList<>();

            // Add all merged regions to the list
            for (int i = 0; i < mergedRegionsCount; i++) {
                CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                if (mergedRegion.getFirstRow() == 0) {
                    mergedRegions.add(mergedRegion);
                }
            }

            // Sort the merged regions by their starting column
            mergedRegions.sort(Comparator.comparingInt(CellRangeAddress::getFirstColumn));

            // Process merged regions in order
            for (int i = 0; i < mergedRegions.size(); i++) {
                CellRangeAddress mergedRegion = mergedRegions.get(i);
                int mergedGroupIndex = i + 1;

                // Get the title of the merged group from row 1
                Cell titleCell = sheet.getRow(0).getCell(mergedRegion.getFirstColumn());
                String title = (titleCell != null) ? titleCell.toString() : "Untitled";

                // Print merged group title
                System.out.println("Merged Group " + mergedGroupIndex + " Title: " + title);

                // Loop through cells in row 2 within the merged region
                Row row2 = sheet.getRow(1); // Row 2
                for (int colIndex = mergedRegion.getFirstColumn(); colIndex <= mergedRegion.getLastColumn(); colIndex++) {
                    // Check if the cell is part of row 2
                    if (colIndex >= 0 && colIndex < row2.getLastCellNum()) {
                        Cell cell = row2.getCell(colIndex);
                        // Print cell value
                        if (cell != null) {
                            System.out.println("Value in column " + (colIndex + 1) + " of row 2: " + cell.toString());

                            // Retrieve and print the value of the cell directly below in row 3
                            Row row3 = sheet.getRow(3); // Row 3
                            if (row3 != null) {
                                Cell cellBelow = row3.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                System.out.println("Value under the cell of row 2 in column " + (colIndex + 1) + ": " + getCellValueAsString(cellBelow));
                            }
                        }
                    }
                }
            }

            System.out.println("Number of merged groups in row 1: " + mergedRegions.size());

            fis.close(); // Close the FileInputStream for the original workbook

            // Open the template workbook
            FileInputStream templateFis = new FileInputStream(new File(templateFilePath));
            Workbook templateWorkbook = new XSSFWorkbook(templateFis);

            // Access the worksheet you want to modify, assuming it's the first one
            Sheet templateSheet = templateWorkbook.getSheetAt(0);

            // Retrieve the style from the template cell A20 (if it exists)
            CellStyle templateStyle = null;
            Row templateRow = templateSheet.getRow(19); // 0-based index for the 4th row
            if (templateRow != null) {
                Cell templateCell = templateRow.getCell(0); // 0-based index for the 1st column
                if (templateCell != null) {
                    templateStyle = templateCell.getCellStyle();
                }
            }

            // Set the value "Grade shs" in cell A20
            int rowIndex = 19; // Rows are 0-based, so this accesses the fourth row
            Row row = templateSheet.getRow(rowIndex);
            if (row == null) {
                row = templateSheet.createRow(rowIndex);
            }
            Cell cell = row.createCell(0); // Columns are 0-based, this accesses the first column (A)
            cell.setCellValue("Grade shs"); // Set the value of the cell

            // Apply the retrieved style to the new cell if it exists
            if (templateStyle != null) {
                cell.setCellStyle(templateStyle);
            }

           /* // Create numbered items from A5 to A30
            for (int i = 21; i <= 48; i++) { // Loop from row index 21 (A21) to 29 (A49)
                row = templateSheet.getRow(i);
                if (row == null) {
                    row = templateSheet.createRow(i);
                }
                cell = row.createCell(0); // Create cell in column A
                cell.setCellValue(i - 20); // Set cell value to the item number

                // Apply the retrieved style to the new cell if it exists
                if (templateStyle != null) {
                    cell.setCellStyle(templateStyle);
                }
            }*/

            // Write the output to a new file
            FileOutputStream fileOut = new FileOutputStream(newFilePath);
            templateWorkbook.write(fileOut);

            // Close all resources
            fileOut.close();
            templateFis.close();
            templateWorkbook.close();

            System.out.println("New Excel file created successfully from the template with 'Grade shs' in cell A4 and numbered items from A5 to A30.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Utility method to retrieve cell value as string
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return ""; // Return empty string for null cells
        }

        DataFormatter formatter = new DataFormatter(); // Creating formatter using the default locale

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return formatter.formatCellValue(cell); // Format date
                } else {
                    return String.valueOf(cell.getNumericCellValue()); // Convert numeric value to string
                }

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula(); // You might want to evaluate formula cells
            case BLANK:
                return ""; // Return empty string for blank cells
            default:
                return "Unknown Cell Type"; // Return this for unknown cell types
        }
    }
}