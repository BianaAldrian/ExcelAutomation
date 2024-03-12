package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelAutomation {
    public static void main(String[] args) {
        // Path to your existing Excel template file
        String templateFilePath = "res/files/itemTemplate.xlsx";

        // Path to save the new Excel file
        String newFilePath = "C:\\Users\\5CG6105SVT\\Desktop\\newFile.xlsx";

        String excelFile = "C:\\Users\\5CG6105SVT\\Desktop\\LTE-SM-2023-Allocation-List-Lots-10111219-Nikka-Trading-1-Copy.xlsx";

        try {
            /*// Open the template workbook
            FileInputStream inputStream = new FileInputStream(new File(templateFilePath));
            Workbook workbook = new XSSFWorkbook(inputStream);

            // Access the worksheet you want to modify, assuming it's the first one
            Sheet sheet = workbook.getSheetAt(0);


            // Retrieve the style from the template cell A4 (if it exists)
            CellStyle templateStyle = null;
            Row templateRow = sheet.getRow(3); // 0-based index for the 4th row
            if (templateRow != null) {
                Cell templateCell = templateRow.getCell(0); // 0-based index for the 1st column
                if (templateCell != null) {
                    templateStyle = templateCell.getCellStyle();
                }
            }

            // Set the value "Grade shs" in cell A4
            int rowIndex = 3; // Rows are 0-based, so this accesses the fourth row
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            Cell cell = row.createCell(0); // Columns are 0-based, this accesses the first column (A)
            cell.setCellValue("Grade shs"); // Set the value of the cell

            // Apply the retrieved style to the new cell if it exists
            if (templateStyle != null) {
                cell.setCellStyle(templateStyle);
            }*/

           /* // Create numbered items from A5 to A30
            for (int i = 4; i <= 29; i++) { // Loop from row index 4 (A5) to 29 (A30)
                row = sheet.getRow(i);
                if (row == null) {
                    row = sheet.createRow(i);
                }
                cell = row.createCell(0); // Create cell in column A
                cell.setCellValue(i - 3); // Set cell value to the item number

                // Apply the retrieved style to the new cell if it exists
                if (templateStyle != null) {
                    cell.setCellStyle(templateStyle);
                }
            }*/

            FileInputStream inputStream = new FileInputStream(new File(excelFile));
            Workbook workbook = new XSSFWorkbook(inputStream);

            // Access the worksheet you want to modify, assuming it's the first one
            Sheet sheet = workbook.getSheetAt(0);

            // Assuming row 1 is the second row (0-based index)
            int rowIndex1 = 0;
            Row row1 = sheet.getRow(rowIndex1);
            if (row1 != null) {
                // Start from column F (5th column, 0-based index)
                int startColumnIndex = 5;
                for (int columnIndex = startColumnIndex; columnIndex < row1.getLastCellNum(); columnIndex++) {
                    Cell cell1 = row1.getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    boolean isMerged = isCellMerged(sheet, cell1);
                    System.out.println("Cell at row " + (rowIndex1 + 1) + ", column " + (columnIndex + 1) +
                            " is " + (isMerged ? "" : "not ") + "merged.");

                    if (isMerged) {
                        String mergedValue = getMergedValue(sheet, cell1);
                        System.out.println("Value of merged group: " + mergedValue);
                    } else {
                        System.out.println("Value of cell: " + cell1.getStringCellValue());
                    }
                }
            }

            // Write the output to a new file
            FileOutputStream fileOut = new FileOutputStream(newFilePath);
            workbook.write(fileOut);

            // Close all resources
            fileOut.close();
            inputStream.close();
            workbook.close();

            System.out.println("New Excel file created successfully from the template with 'Grade shs' in cell A4 and numbered items from A5 to A30.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean isCellMerged(Sheet sheet, Cell cell) {
        int cellRow = cell.getRowIndex();
        int cellColumn = cell.getColumnIndex();

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.isInRange(cellRow, cellColumn)) {
                return true;
            }
        }
        return false;
    }

    private static String getMergedValue(Sheet sheet, Cell cell) {
        int cellRow = cell.getRowIndex();
        int cellColumn = cell.getColumnIndex();

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.isInRange(cellRow, cellColumn)) {
                Row firstRow = sheet.getRow(range.getFirstRow());
                Cell firstCell = firstRow.getCell(range.getFirstColumn());
                return firstCell.getStringCellValue(); // Return the value of the top-left cell in the merged region
            }
        }
        return null; // Return null if the cell is not part of a merged region
    }
}