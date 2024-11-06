package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ExcelCellColorReader {

    public void printCellColors() {
        String filePath = "F:\\My learning items\\JMeter Data Preparation\\test_new.xlsx"; // Replace with your file path

        try (InputStream inp = new FileInputStream(filePath)) {
            Workbook workbook;
            if (filePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(inp);
            } else {
                workbook = new HSSFWorkbook(inp);
            }

            Sheet sheet = workbook.getSheetAt(0); // Get first sheet

            // Iterate over the first 4 rows
            for (int rowIndex = 1; rowIndex < 3; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue; // Skip if row is empty

                // Iterate over all cells in the current row
                for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    if (cell == null) continue; // Skip if cell is empty

                    CellStyle style = cell.getCellStyle();

                    System.out.print("Row " + (rowIndex + 1) + ", Cell " + (cellIndex + 1) + ": ");

                    // Check for background color
                    if (style.getFillPattern() != FillPatternType.NO_FILL) {
                        if (style.getFillForegroundColorColor() != null) {
                            Color bgColor = style.getFillForegroundColorColor();
                            System.out.print("Background color: " + bgColor + " ");
                        }
                    } else {
                        System.out.print("No background color ");
                    }

                    // Check for font color
                    Font font = workbook.getFontAt(style.getFontIndexAsInt());
                    if (font.getColor() != Font.COLOR_NORMAL) {
                        System.out.print("Font color: " + font.getColor());
                    } else {
                        System.out.print("No font color");
                    }

                    System.out.println(); // Move to the next line for the next cell
                }
            }

            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
