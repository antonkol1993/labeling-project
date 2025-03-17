package com.anton.labeling.service.largeBox;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class CelFiller {

    private final XSSFWorkbook workbook;
    private final XSSFSheet sheet;

    public CelFiller(XSSFWorkbook workbook, XSSFSheet sheet) {
        this.workbook = workbook;
        this.sheet = sheet;
    }

    public void fillCells(ItemLargeBox item, int startRow, int startCol) {
        Font arialFontMain = createArialFont((short) 11);
        setCellValueWithStyle(startRow + 3, startCol, item.getName() + "\n" + item.getSize(), arialFontMain, HorizontalAlignment.CENTER);

        Font arialFontDownRows = createArialFont((short) 10);
        for (int rowNum = 4; rowNum <= 10; rowNum++) {
            if (rowNum == 9) continue;
            int adjustedRow = startRow + rowNum;
            if (rowNum == 4) {
                setCellValueWithStyle(adjustedRow, startCol, "Marking", arialFontDownRows, HorizontalAlignment.LEFT);
                setCellValueWithStyle(adjustedRow, startCol + 1, item.getMarking(), arialFontDownRows, HorizontalAlignment.CENTER);
            }
            if (rowNum == 5) {
                setCellValueWithStyle(adjustedRow, startCol, "РАЗМЕР/Size", arialFontDownRows, HorizontalAlignment.LEFT);
            }
        }

        Font arialFontCenter = createArialFont((short) 10);
        setCellValueWithMergedStyle(startRow + 5, startCol + 1, startRow + 5, startCol + 2, item.getSize(), arialFontCenter, HorizontalAlignment.CENTER);
    }

    private Font createArialFont(short fontSize) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        font.setFontHeightInPoints(fontSize);
        return font;
    }

    private void setCellValueWithStyle(int rowNum, int colNum, String value, Font font, HorizontalAlignment alignment) {
        Row row = sheet.getRow(rowNum);
        if (row == null) row = sheet.createRow(rowNum);

        Cell cell = row.getCell(colNum);
        if (cell == null) cell = row.createCell(colNum);

        cell.setCellValue(value);

        CellStyle style = CardStyle.createBorderedCellStyle(workbook, rowNum, colNum, 0, 0);
        style.setFont(font);
        style.setAlignment(alignment);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        cell.setCellStyle(style);
    }

    private void setCellValueWithMergedStyle(int startRow, int startCol, int endRow, int endCol, String value, Font font, HorizontalAlignment alignment) {
        setCellValueWithStyle(endRow, startCol, value, font, alignment);
        if (!isMergedRegion(startRow, startCol, endRow, endCol)) {
            sheet.addMergedRegion(new CellRangeAddress(startRow, startCol, endRow, endCol));
        }
    }

    private boolean isMergedRegion(int startRow, int startCol, int endRow, int endCol) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstRow() == startRow && range.getFirstColumn() == startCol &&
                    range.getLastRow() == endRow && range.getLastColumn() == endCol) {
                return true;
            }
        }
        return false;
    }
}
