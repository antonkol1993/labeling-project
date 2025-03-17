package com.anton.labeling.service.largeBox;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellFiller {
    private final XSSFWorkbook workbook;
    private final XSSFSheet sheet;

    public CellFiller(XSSFWorkbook workbook, XSSFSheet sheet) {
        this.workbook = workbook;
        this.sheet = sheet;
    }

    public void fillCellsWithThree(int startRow, int startCol, int endRow, int endCol) {
        for (int row = startRow; row <= endRow; row++) {
            Row sheetRow = sheet.getRow(row);
            if (sheetRow == null) {
                sheetRow = sheet.createRow(row);
            }
            for (int col = startCol; col <= endCol; col++) {
                Cell cell = sheetRow.getCell(col);
                if (cell == null) {
                    cell = sheetRow.createCell(col);
                }
                cell.setCellValue("3");

                // Применяем стиль из CardStyle
                CellStyle cellStyle = CardStyle.createBorderedCellStyle(workbook, startRow, startCol, row, col);
                applyBoldCenteredStyle(cellStyle);
                cell.setCellStyle(cellStyle);
            }
        }
    }

    private void applyBoldCenteredStyle(CellStyle cellStyle) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 11);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    }
}