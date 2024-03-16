package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelAutomation {

    private static final String HEADER = "res/images/header.png";
    private static final String EXCEL_FILE_PATH = "C:\\Users\\5CG6105SVT\\Desktop\\LTE-SM-2023-Allocation-List-Lots-10111219-Nikka-Trading-1-Copy.xlsx";
    private static final String TEMPLATE_FILE_PATH = "res/files/receiptTemplate.xlsx";
    private static final String NEW_FILE_PATH = "C:\\Users\\5CG6105SVT\\Desktop\\newFile.xlsx";

    private static final int COLUMN_INDEX_C = 2;

    private static String region, division, schoolID, schoolName, gradeLevel;

    public static void main(String[] args) {
        try (
                FileInputStream fis = new FileInputStream(EXCEL_FILE_PATH);
                Workbook workbook = WorkbookFactory.create(fis);
                FileInputStream templateFis = new FileInputStream(new File(TEMPLATE_FILE_PATH));
                Workbook templateWorkbook = new XSSFWorkbook(templateFis);
                FileOutputStream fileOut = new FileOutputStream(NEW_FILE_PATH)
        ) {
            Sheet sheet = workbook.getSheetAt(0);

            Map<String, StringBuilder> groupsC = new HashMap<>();
            for (int rowIndex = 3; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cellC = row.getCell(COLUMN_INDEX_C);
                if (cellC != null) {
                    String cellValueC = getCellValueAsString(cellC);
                    groupsC.computeIfAbsent(cellValueC, k -> new StringBuilder()).append(rowIndex).append(",");
                }
            }

            for (Map.Entry<String, StringBuilder> entry : groupsC.entrySet()) {
                String[] rowIndices = entry.getValue().toString().split(",");
                int[] rowNumbers = Arrays.stream(rowIndices).mapToInt(Integer::parseInt).toArray();
                for (int rowNum : rowNumbers) {
                    Row row = sheet.getRow(rowNum);
                    for (int colIndex = 0; colIndex <= 4; colIndex++) {
                        Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        String cellValue = getCellValueAsString(cell);
                        switch (colIndex) {
                            case 0 -> region = cellValue;
                            case 1 -> division = cellValue;
                            case 2 -> schoolID = cellValue;
                            case 3 -> schoolName = cellValue;
                            case 4 -> gradeLevel = cellValue;
                        }
                    }

                    List<CellRangeAddress> mergedRegions = getMergedRegions(sheet);
                    for (int i = 0; i < 1; i++) {
                        CellRangeAddress mergedRegion = mergedRegions.get(i);
                        processMergedRegion(templateWorkbook, sheet, mergedRegion, i + 1, row);
                    }

                    updateTemplate(templateWorkbook, 11, 1, schoolName);
                    updateTemplate(templateWorkbook, 11, 3, schoolID);
                    updateTemplate(templateWorkbook, 12, 1, division);
                    updateTemplate(templateWorkbook, 12, 3, region);
                    updateTemplate(templateWorkbook, 19, 0, gradeLevel);

                }
                break;
            }

            templateWorkbook.write(fileOut);
            System.out.println("New Excel file created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return formatter.formatCellValue(cell);
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "Unknown Cell Type";
        }
    }

    private static void addHeaderImage(Sheet sheet, Workbook workbook, int newRowStart) {
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper helper = workbook.getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(0);
        anchor.setCol2(3);
        int dx2 = Units.toEMU(10 * Units.DEFAULT_CHARACTER_WIDTH);
        anchor.setDx2(dx2);
        anchor.setRow1(newRowStart + 1);
        anchor.setRow2(newRowStart + 6);
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

        int columnIndex = sourceCell.getColumnIndex();
        int width = sheet.getColumnWidth(columnIndex);
        sheet.setColumnWidth(newCell.getColumnIndex(), width);

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

    private static void processMergedRegion(Workbook templateWorkbook, Sheet sheet, CellRangeAddress mergedRegion, int mergedGroupIndex, Row row) {
        // Extracting title from the first row of the merged region
        Cell titleCell = sheet.getRow(0).getCell(mergedRegion.getFirstColumn());
        String title = (titleCell != null) ? titleCell.toString() : "Untitled";
        System.out.println("Merged Group " + mergedGroupIndex + " Title: " + title);
        // Updating the template with the extracted title
        updateTemplate(templateWorkbook, 17, 1, title);

        // Initializing variables
        int startRow = 20;
        int rowCount = 0; // Counter for row count
        Row row2 = sheet.getRow(1);

        // Iterating over the cells within the merged region
        for (int colIndex = mergedRegion.getFirstColumn(); colIndex <= mergedRegion.getLastColumn(); colIndex++) {
            if (colIndex >= 0 && colIndex < row2.getLastCellNum()) {
                Cell cell = row2.getCell(colIndex);
                if (cell != null) {
                    // Updating template with cell value
                    updateTemplate(templateWorkbook, startRow, 1, cell.toString());
                    if (row != null) {
                        Cell cellBelow = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        System.out.println("Item Name: " + cell.toString() + " QTY: " + getCellValueAsString(cellBelow));
                        // Updating template with cell value below
                        updateTemplate(templateWorkbook, startRow, 2, getCellValueAsString(cellBelow));
                    }
                    startRow++;
                    rowCount++;
                }

                // Checking if rowCount reaches 24
                if (rowCount == 24) {
                    // Create a new template when rowCount reaches 24

                    // Get the template sheet and its last row number
                    Sheet templateSheet = templateWorkbook.getSheetAt(0);
                    int lastRowNum = templateSheet.getLastRowNum();

                    // Define the starting row for the new content
                    int newRowStart = (lastRowNum + 3);

                    // Add header image to the new content
                    addHeaderImage(templateSheet, templateWorkbook, newRowStart);

                    // Clone rows from the template to the new content
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
                    break; // Exit the loop once the new content is created
                }
            }
        }
    }



    private static void updateTemplate(Workbook templateWorkbook, int rowNum, int colNum, String value) {
        Sheet templateSheet = templateWorkbook.getSheetAt(0);
        CellStyle templateStyle = getTemplateCellStyle(templateSheet, rowNum, colNum);
        Row row = templateSheet.getRow(rowNum);
        if (row == null) {
            row = templateSheet.createRow(rowNum);
        }
        Cell cell = row.createCell(colNum);
        cell.setCellValue(value);
        if (templateStyle != null) {
            cell.setCellStyle(templateStyle);
        }
    }

    private static CellStyle getTemplateCellStyle(Sheet templateSheet, int rowNum, int colNum) {
        Row templateRow = templateSheet.getRow(rowNum);
        if (templateRow != null) {
            Cell templateCell = templateRow.getCell(colNum);
            if (templateCell != null) {
                return templateCell.getCellStyle();
            }
        }
        return null;
    }

    private static List<CellRangeAddress> getMergedRegions(Sheet sheet) {
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.getFirstRow() == 0) {
                mergedRegions.add(mergedRegion);
            }
        }
        mergedRegions.sort(Comparator.comparingInt(CellRangeAddress::getFirstColumn));
        return mergedRegions;
    }
}
