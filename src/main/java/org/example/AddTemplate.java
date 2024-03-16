package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class AddTemplate {

    private static final String HEADER = "res/images/header.png";

    /*private static final String HEADER = "res/images/header.png";
    private static final String TEMPLATE_FILE_PATH = "res/files/receiptTemplate.xlsx";
    private static final String NEW_FILE_PATH = "C:\\Users\\5CG6105SVT\\Desktop\\newFile1.xlsx";*/

    /*public static void main(String[] args) {
        try {
            FileInputStream fis = new FileInputStream(TEMPLATE_FILE_PATH);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet templateSheet = workbook.getSheetAt(0); // Get the first sheet from the workbook
            int lastRowNum = templateSheet.getLastRowNum();

            // Let's say you want to create 5 copies of the template in the same sheet
            for (int i = 1; i < 2 ; i++) {
                int newRowStart = (lastRowNum + 3) * i; // Calculate the starting row index for the new copy
                // Add the header image after copying each template
                addHeaderImage(templateSheet, workbook, newRowStart);
                for (int j = 0; j <= lastRowNum; j++) {
                    Row sourceRow = templateSheet.getRow(j);
                    Row newRow = templateSheet.createRow(newRowStart + j);
                    if (sourceRow != null) {
                        for (int k = 0; k < sourceRow.getLastCellNum(); k++) {
                            Cell sourceCell = sourceRow.getCell(k);
                            if (sourceCell != null) {
                                Cell newCell = newRow.createCell(k);
                                cloneCell(sourceCell, newCell, templateSheet);
                            }
                        }
                    }
                }

            }

            // Write the modified workbook to a new file
            try (FileOutputStream fos = new FileOutputStream(NEW_FILE_PATH)) {
                workbook.write(fos);
            }

            System.out.println("Template has been copied multiple times to: " + NEW_FILE_PATH);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/

    public void addTemplate(Workbook workbook, String outputFile) {
        int newRowStart = 0;
        try {
            Sheet templateSheet = workbook.getSheetAt(0); // Get the first sheet from the workbook
            int lastRowNum = templateSheet.getLastRowNum();

            // Calculate the starting row index for the new copy
            newRowStart = (lastRowNum + 3); // Assuming only one copy is added

            // Add the header image after copying the template
            addHeaderImage(templateSheet, workbook, newRowStart);

            // Copy the template
            for (int j = 0; j <= lastRowNum; j++) {
                Row sourceRow = templateSheet.getRow(j);
                Row newRow = templateSheet.createRow(newRowStart + j);
                if (sourceRow != null) {
                    for (int k = 0; k < sourceRow.getLastCellNum(); k++) {
                        Cell sourceCell = sourceRow.getCell(k);
                        if (sourceCell != null) {
                            Cell newCell = newRow.createCell(k);
                            cloneCell(sourceCell, newCell, templateSheet);
                        }
                    }
                }
            }

            // Write the modified workbook to a new file
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            System.out.println("Template has been copied to: " + outputFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static void addHeaderImage(Sheet sheet, Workbook workbook, int newRowStart) {
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper helper = workbook.getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(0); // Column 0 (Column A)
        anchor.setCol2(3); // Column 3 (Column D)

        // Calculate the X coordinate within column D for the fractional value (e.g., 0.9 for 3.9)
        int dx2 = Units.toEMU(10 * Units.DEFAULT_CHARACTER_WIDTH); // Assuming 0.9 is the fractional part
        anchor.setDx2(dx2);

        anchor.setRow1(newRowStart + 1); // Adjusted for the new row start
        anchor.setRow2(newRowStart + 6); // Adjusted for the new row end (adjust as needed)

        try (FileInputStream fis = new FileInputStream(HEADER)) {
            byte[] bytes = fis.readAllBytes();
            int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
            drawing.createPicture(anchor, pictureIndex);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void cloneCell(Cell sourceCell, Cell newCell, Sheet sheet) {
        newCell.setCellStyle(sourceCell.getCellStyle());
        switch (sourceCell.getCellType()) {
            case STRING:
                newCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                newCell.setCellType(CellType.BLANK);
                break;
            case ERROR:
                newCell.setCellErrorValue(sourceCell.getErrorCellValue());
                break;
            default:
                newCell.setCellValue(sourceCell.toString());
                break;
        }

        // Copy the column width from the source column to the new column
        int columnIndex = sourceCell.getColumnIndex();
        int width = sheet.getColumnWidth(columnIndex);
        sheet.setColumnWidth(newCell.getColumnIndex(), width);

        // Check if the source cell is part of a merged region and if so, copy the merged region to the new cell
        for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
            if (mergedRegion.isInRange(sourceCell.getRowIndex(), sourceCell.getColumnIndex())) {
                CellRangeAddress newMergedRegion = new CellRangeAddress(
                        newCell.getRowIndex(),
                        newCell.getRowIndex() + (mergedRegion.getLastRow() - mergedRegion.getFirstRow()),
                        newCell.getColumnIndex(),
                        newCell.getColumnIndex() + (mergedRegion.getLastColumn() - mergedRegion.getFirstColumn())
                );
                if (!isOverlapping(newMergedRegion, sheet.getMergedRegions())) {
                    sheet.addMergedRegion(newMergedRegion);
                }
                break;
            }
        }
    }

    private static boolean isOverlapping(CellRangeAddress newRegion, List<CellRangeAddress> existingRegions) {
        for (CellRangeAddress existingRegion : existingRegions) {
            if (newRegion.isInRange(existingRegion.getFirstRow(), existingRegion.getFirstColumn()) ||
                    newRegion.isInRange(existingRegion.getLastRow(), existingRegion.getLastColumn())) {
                return true;
            }
        }
        return false;
    }

}
