package com.anton.labeling.service.irrelevantOld;

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

    public void fillCells(XSSFWorkbook workbook, XSSFSheet sheet, ItemLargeBox item) {

        // 🔹 Жирный шрифт для заголовков
        Font arialFontMain = createArialFont((short) 11);



// Заполнение B3 (жирный центр)
        String nameAndSize = item.getName() + "\n" + item.getSize();
        setCellValueWithStyle(3, 1, nameAndSize, arialFontMain, HorizontalAlignment.CENTER);
// Убедитесь, что ячейка имеет перенос текста
        Row row3 = sheet.getRow(3);
        if (row3 == null) row3 = sheet.createRow(3);
        Cell cell3 = row3.getCell(1);
        if (cell3 == null) cell3 = row3.createCell(1);
        CellStyle style = cell3.getCellStyle();
        style.setWrapText(true);  // Включаем перенос текста
        cell3.setCellStyle(style);
// Автоматическая настройка ширины столбца для корректного отображения
        sheet.autoSizeColumn(1);  // Индекс столбца B
// Установите высоту строки, чтобы все текстовые строки поместились
        row3.setHeightInPoints(50);  // Примерная высота, вы можете настроить по своему усмотрению

        // 🔹 Жирный шрифт для B4-B10 (кроме B9) с левым выравниванием
        Font arialFontDownRows = createArialFont((short) 10);
        for (int rowNum = 4; rowNum <= 10; rowNum++) {
            if (rowNum == 9) continue; // Пропускаем строку 9
            if (rowNum == 4) {
                setCellValueWithStyle(rowNum, 1, "Marking", arialFontDownRows, HorizontalAlignment.LEFT);
                setCellValueWithStyle(rowNum, 2, item.getMarking(), arialFontDownRows, HorizontalAlignment.CENTER);
            }
            if (rowNum == 5) {
                setCellValueWithStyle(rowNum, 1, "РАЗМЕР/Size", arialFontDownRows, HorizontalAlignment.LEFT);
            }
            if (rowNum == 7) {
                setCellValueWithStyle(rowNum, 1, "Кол-во в упак/шт.", arialFontDownRows, HorizontalAlignment.LEFT);
            }
            if (rowNum == 8) {
                setCellValueWithStyle(rowNum, 1, "Вес упак Кг/Kgs", arialFontDownRows, HorizontalAlignment.LEFT);
            }
            if (rowNum == 10) {
                setCellValueWithStyle(rowNum, 1, "ORDER:", arialFontDownRows, HorizontalAlignment.LEFT);
            }

        }

        // 🔹 Жирный шрифт для C5:D5, C7 (центр) + item
        Font arialFontCenter = createArialFont((short) 10);
        setCellValueWithMergedStyle(5, 2, item.getSize(), arialFontCenter, HorizontalAlignment.CENTER, 5, 3);
        setCellValueWithStyle(7, 2, item.getQuantityInBox(), arialFontCenter, HorizontalAlignment.CENTER);

        // 🔹 Заполнение D7-D8, C9-C10 (текст слева, жирный)
        Font arialFontFinalLeft = createArialFont((short) 10);
        for (int rowNum = 7; rowNum <= 10; rowNum++) {
            if (rowNum == 7) {
                setCellValueWithStyle(rowNum, 3, "Шт / PCS", arialFontFinalLeft, HorizontalAlignment.LEFT);
            }
            if (rowNum == 8) {
                setCellValueWithStyle(rowNum, 3, "Кг/Kgs", arialFontFinalLeft, HorizontalAlignment.LEFT);
            }

            if (rowNum == 9) {
                setCellValueWithStyle(rowNum, 2, "Сделано в КНР", arialFontFinalLeft, HorizontalAlignment.LEFT);
            }
            if (rowNum == 10) {
                setCellValueWithStyle(rowNum, 2, item.getOrder(), arialFontFinalLeft, HorizontalAlignment.LEFT);
            }
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
    private void setCellValueWithMergedStyle(int endRow, int startCol, String value, Font font, HorizontalAlignment alignment,
                                             int startRow, int endCol) {
        setCellValueWithStyle(endRow, startCol, value, font, alignment);

        if (!isMergedRegion(startRow, startCol, endRow, endCol)) {
            sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, startCol, endCol));
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