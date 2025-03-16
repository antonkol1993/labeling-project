package com.anton.labeling.service.largeBox;
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

    // Создание шрифта Arial
    public Font createArialFont(short fontSize) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        font.setFontHeightInPoints(fontSize);
        return font;
    }

    // Устанавливает значение и стиль ячейки с заданным выравниванием
    public void setCellValueWithStyle(int rowNum, int colNum, String value, Font font, HorizontalAlignment alignment) {
        Row row = sheet.getRow(rowNum);
        if (row == null) row = sheet.createRow(rowNum);

        Cell cell = row.getCell(colNum);
        if (cell == null) cell = row.createCell(colNum);

        cell.setCellValue(value);

        CellStyle style = CardStyle.createBorderedCellStyle(workbook, sheet, rowNum, colNum);
        style.setFont(font);
        style.setAlignment(alignment);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        cell.setCellStyle(style);
    }

    // Устанавливает значение в ячейку с объединением ячеек и заданным стилем
    public void setCellValueWithMergedStyle(int rowNum, int colNum, String value, Font font, HorizontalAlignment alignment, int mergeStartRow, int mergeEndCol) {
        setCellValueWithStyle(rowNum, colNum, value, font, alignment);

        if (!isMergedRegion(mergeStartRow, colNum, rowNum, mergeEndCol)) {
            sheet.addMergedRegion(new CellRangeAddress(mergeStartRow, rowNum, colNum, mergeEndCol));
        }
    }

    // Проверяет, существует ли уже объединенный регион в заданном диапазоне
    private boolean isMergedRegion(int firstRow, int firstCol, int lastRow, int lastCol) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstRow() == firstRow && range.getFirstColumn() == firstCol &&
                    range.getLastRow() == lastRow && range.getLastColumn() == lastCol) {
                return true;
            }
        }
        return false;
    }
}