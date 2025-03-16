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

    public void fillCells(XSSFWorkbook workbook, XSSFSheet sheet) {
        CelFiller cellFiller = new CelFiller(workbook, sheet);

        // 🔹 Жирный шрифт для заголовков
        Font arialFontMain = cellFiller.createArialFont((short) 11);

        // Заполнение B3 (жирный центр)
        cellFiller.setCellValueWithStyle(3, 1, "Жирный Центр Item", arialFontMain, HorizontalAlignment.CENTER);

        // 🔹 Жирный шрифт для B4-B10 (кроме B9) с левым выравниванием
        Font arialFontDownRows = cellFiller.createArialFont((short) 10);
        for (int rowNum = 4; rowNum <= 10; rowNum++) {
            if (rowNum == 9) continue; // Пропускаем строку 9
            cellFiller.setCellValueWithStyle(rowNum, 1, "Жирный лево Item", arialFontDownRows, HorizontalAlignment.LEFT);
        }

        // 🔹 Жирный шрифт для C5:D5, C7 (центр)
        Font arialFontCenter = cellFiller.createArialFont((short) 10);
        cellFiller.setCellValueWithMergedStyle(5, 2, "Жирный Центр Item", arialFontCenter, HorizontalAlignment.CENTER, 5, 3);
        cellFiller.setCellValueWithStyle(7, 2, "Жирный Центр Item", arialFontCenter, HorizontalAlignment.CENTER);

        // 🔹 Заполнение D6-D8, C9-C10 (текст слева, жирный)
        Font arialFontFinalLeft = cellFiller.createArialFont((short) 10);
        for (int rowNum = 6; rowNum <= 8; rowNum++) {
            cellFiller.setCellValueWithStyle(rowNum, 3, "final left", arialFontFinalLeft, HorizontalAlignment.LEFT);
        }
        for (int rowNum = 9; rowNum <= 10; rowNum++) {
            cellFiller.setCellValueWithStyle(rowNum, 2, "final left", arialFontFinalLeft, HorizontalAlignment.LEFT);
        }
    }

    // Создание шрифта Arial
    private Font createArialFont(short fontSize) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        font.setFontHeightInPoints(fontSize);
        return font;
    }

    // Устанавливает значение и стиль ячейки с заданным выравниванием
    private void setCellValueWithStyle(int rowNum, int colNum, String value, Font font, HorizontalAlignment alignment) {
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
    private void setCellValueWithMergedStyle(int rowNum, int colNum, String value, Font font, HorizontalAlignment alignment, int mergeStartRow, int mergeEndCol) {
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