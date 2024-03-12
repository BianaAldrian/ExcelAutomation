package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.IOUtils;

import java.io.*;

public class ExcelWriter {
    public static void main(String[] args) {
        // Path to your image file in the res directory
        String imagePath = "res/images/header.png"; // Adjust the file name and path accordingly

        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        try {
            // Load the image from the specified path
            InputStream inputStream = new FileInputStream(imagePath);
            byte[] bytes = IOUtils.toByteArray(inputStream);
            int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
            inputStream.close();

            // Set the page layout size to legal (21.59 x 35.56 cm)
            sheet.getPrintSetup().setPaperSize(PrintSetup.LEGAL_PAPERSIZE);
            sheet.getPrintSetup().setLandscape(false); // Set to true if you want landscape orientation

            // Set the print margins
            sheet.setMargin(Sheet.LeftMargin, 0.6); // Set the left margin to 0.6 inches
            sheet.setMargin(Sheet.TopMargin, 1.3); // Set the top margin to 1.3 inches
            sheet.setMargin(Sheet.RightMargin, 0.6); // Set the right margin to 0.6 inches
            sheet.setMargin(Sheet.BottomMargin, 1.3); // Set the bottom margin to 1.3 inches

            // Set header and footer margins
            sheet.setMargin(Sheet.HeaderMargin, 0.8); // Set the header margin to 0.8 inches
            sheet.setMargin(Sheet.FooterMargin, 0.8); // Set the footer margin to 0.8 inches

            // Center on page horizontally
            sheet.setHorizontallyCenter(true);

            // Create a drawing canvas
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            // Set the width of column A to approximately 80 pixels
            final int PIXELS_PER_CHAR = 7; // Approximate pixel width of one character
            final int MAX_DIGIT_WIDTH = 256; // Excel's maximum digit width unit
            int colWidth = (int) ((80.0 / PIXELS_PER_CHAR) * MAX_DIGIT_WIDTH);
            sheet.setColumnWidth(0, colWidth); // Set the width of column A
            sheet.setColumnWidth(6, colWidth);

            // Create an anchor for the image
            CreationHelper helper = workbook.getCreationHelper();
            ClientAnchor anchor = helper.createClientAnchor();

            // Set top-left corner of the image:
            anchor.setCol1(0); // Column A
            anchor.setRow1(0); // Row 1

            // Set bottom-right corner of the image:
            anchor.setCol2(9); // Column I (0-based index, 9 means up to the start of column J)
            anchor.setRow2(5); // Adjust row height as needed

            // Create the picture
            Picture picture = drawing.createPicture(anchor, pictureIndex);

            // Add the title below the image
            int titleRowIndex = anchor.getRow2() + 1; // This will be the row just below the image
            Row titleRow = sheet.createRow(titleRowIndex);
            Cell titleCell = titleRow.createCell(0); // Create a cell in the first column
            titleCell.setCellValue("DELIVERY RECEIPT (DEPED)"); // Set the title text

            // Create a font and style for the title
            Font titleFont = workbook.createFont();
            titleFont.setFontName("Imprint MT Shadow");
            titleFont.setFontHeightInPoints((short) 18);
            titleFont.setItalic(true);
            titleFont.setBold(true);
            titleFont.setColor(IndexedColors.BLACK.getIndex()); // Set the font color to black

            CellStyle titleStyle = workbook.createCellStyle();
            titleStyle.setAlignment(HorizontalAlignment.CENTER);
            titleStyle.setFont(titleFont);
            titleCell.setCellStyle(titleStyle);

            // Merge the title cells
            sheet.addMergedRegion(new CellRangeAddress(
                    titleRowIndex, // First row (0-based)
                    titleRowIndex, // Last row  (0-based)
                    0,             // First column (0-based)
                    8              // Last column  (0-based)
            ));

            // Add the additional text below the title
            int additionalTextRowIndex = titleRowIndex + 1; // The next row after the title
            Row additionalTextRow = sheet.createRow(additionalTextRowIndex);
            additionalTextRow.setHeightInPoints(35); // Set the row height to 35 points (approximately 35 pixels)
            Cell additionalTextCell = additionalTextRow.createCell(8); // Create a cell in column I
            additionalTextCell.setCellValue("No : 2024-000-025"); // Set the additional text

            // Create a font and style for the additional text
            Font additionalTextFont = workbook.createFont();
            additionalTextFont.setFontName("Times New Roman");
            additionalTextFont.setFontHeightInPoints((short) 14);
            additionalTextFont.setBold(true);
            additionalTextFont.setColor(IndexedColors.RED.getIndex()); // Set the font color to red

            CellStyle additionalTextStyle = workbook.createCellStyle();
            additionalTextStyle.setAlignment(HorizontalAlignment.RIGHT);
            additionalTextStyle.setVerticalAlignment(VerticalAlignment.CENTER); // Set vertical alignment to center
            additionalTextStyle.setFont(additionalTextFont);
            additionalTextCell.setCellStyle(additionalTextStyle);

            // Add the text "Delivered to:" in cell A9
            int deliveryInfoRowIndex = 8; // Row index for A9 (0-based)
            Row deliveryInfoRow = sheet.createRow(deliveryInfoRowIndex);
            deliveryInfoRow.setHeightInPoints(27);
            Cell deliveryInfoCell = deliveryInfoRow.createCell(0); // Cell A9
            deliveryInfoCell.setCellValue("Delivered to:"); // Set the text

            // Create a font and style for the "Delivered to:" text
            Font deliveryInfoFont = workbook.createFont();
            deliveryInfoFont.setFontName("Gisha");
            deliveryInfoFont.setFontHeightInPoints((short) 9);
            deliveryInfoFont.setColor(IndexedColors.BLACK.getIndex()); // Set the font color to black

            CellStyle deliveryInfoStyle = workbook.createCellStyle();
            deliveryInfoStyle.setFont(deliveryInfoFont);
            deliveryInfoCell.setCellStyle(deliveryInfoStyle);

            // Create a CellStyle for the merged region with a bottom border
            CellStyle mergedCellStyle = workbook.createCellStyle();
            mergedCellStyle.setBorderBottom(BorderStyle.THIN); // Set a thin bottom border

            // Apply the style to each cell in the merged region
            for (int colIndex = 1; colIndex <= 5; colIndex++) {
                Cell cell = sheet.getRow(8).createCell(colIndex);
                cell.setCellStyle(mergedCellStyle);
            }

            // Merge the title cells and apply the bottom border
            sheet.addMergedRegion(new CellRangeAddress(
                    8, // First row (0-based)
                    8, // Last row  (0-based)
                    1, // First column (0-based)
                    5  // Last column  (0-based)
            ));

            // Ensure that the border is drawn by setting the style again after merging
            RegionUtil.setBorderBottom(BorderStyle.HAIR, new CellRangeAddress(8, 8, 1, 5), sheet);

            // Save the workbook to a file
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\5CG6105SVT\\Desktop\\custom_layout.xlsx"); // Output file name
            workbook.write(fileOut);
            fileOut.close();

            System.out.println("Excel file with image, styled title, and adjusted column width created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}