package com.anton.labeling.service.largeBox;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CardStyle {

    // Метод для создания стиля ячейки
    public static CellStyle createBorderedCellStyle(XSSFWorkbook workbook, int startRow, int startCol, int endRow, int endCol) {
        CellStyle cellStyle = workbook.createCellStyle();
        int relativeRow = endRow - startRow + 1; // Приводим к диапазону 1-10
        int relativeCol = endCol - startCol + 1; // Приводим к диапазону 1-3

        // Для строк 1-3 — границы полностью Medium
        if (relativeRow <= 3) {
            cellStyle.setBorderTop(BorderStyle.MEDIUM);
            cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            cellStyle.setBorderLeft(BorderStyle.MEDIUM);
            cellStyle.setBorderRight(BorderStyle.MEDIUM);
        }
        // Строка 4 (Medium по бокам, остальное Thin)
        else if (relativeRow == 4) {
            if (relativeCol == 1) cellStyle.setBorderLeft(BorderStyle.MEDIUM);
            if (relativeCol == 3) cellStyle.setBorderRight(BorderStyle.MEDIUM);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            if (relativeCol == 1) cellStyle.setBorderRight(BorderStyle.THIN);
        }
        // Настройка остальных строк
        else {
            // Левая граница (1-й столбец)
            if (relativeCol == 1) {
                cellStyle.setBorderLeft(BorderStyle.MEDIUM);
                if (relativeRow >= 5 && relativeRow <= 10) cellStyle.setBorderRight(BorderStyle.THIN);
            }

            // Средний столбец (2-й)
            if (relativeCol == 2) {
                if (relativeRow >= 6 && relativeRow <= 8) {
                    cellStyle.setBorderRight(BorderStyle.THIN); // 6C-8C (без merge)
                }
            }

            // Правая граница (3-й столбец)
            if (relativeCol == 3) {
                cellStyle.setBorderRight(BorderStyle.MEDIUM);
            }

            // Нижняя граница
            if (relativeRow < 10) {
                cellStyle.setBorderBottom(BorderStyle.THIN);
            } else if (relativeRow == 10) {
                cellStyle.setBorderBottom(BorderStyle.MEDIUM); // Вся 10-я строка
            }
        }

        return cellStyle;
    }

    public static CellStyle createBoldCellStyle(Workbook workbook, HorizontalAlignment alignment, short fontSize) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints(fontSize);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(alignment);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        return style;
    }

    public static CellStyle createCenteredBoldStyle(Workbook workbook, short fontSize) {
        return createBoldCellStyle(workbook, HorizontalAlignment.CENTER, fontSize);
    }

    public static CellStyle createLeftBoldStyle(Workbook workbook, short fontSize) {
        return createBoldCellStyle(workbook, HorizontalAlignment.LEFT, fontSize);
    }

    // Метод для добавления объединенных ячеек БЕЗ проверки
    public static void addMergedRegions(XSSFSheet sheet, int startRow, int startCol) {
        int[][] mergedRegions = {
                {startRow, startRow, startCol, startCol + 2}, // 1 строка
                {startRow + 1, startRow + 1, startCol, startCol + 2}, // 2 строка
                {startRow + 2, startRow + 2, startCol, startCol + 2}, // 3 строка
                {startRow + 3, startRow + 3, startCol + 1, startCol + 2}, // 4 строка
                {startRow + 4, startRow + 4, startCol + 1, startCol + 2}, // 5 строка (C5:D5)
                {startRow + 8, startRow + 8, startCol + 1, startCol + 2},  // 9 строка
                {startRow + 9, startRow + 9, startCol + 1, startCol + 2}   // 10 строка
        };

        // Добавляем объединенные области БЕЗ проверки
        for (int[] region : mergedRegions) {
            sheet.addMergedRegion(new CellRangeAddress(region[0], region[1], region[2], region[3]));
        }
    }

    // Новый метод для установки ширины столбцов
    public static void setColumnWidths(XSSFSheet sheet, int startCol) {
        sheet.setColumnWidth(startCol, (int) (((124 - 5) / 7.0 + 0.71) * 256)); // 1-й столбец (124 px)
        sheet.setColumnWidth(startCol + 1, (int) (((88 - 5) / 7.0 + 0.71) * 256)); // 2-й столбец (88 px)
        sheet.setColumnWidth(startCol + 2, (int) (((88 - 5) / 7.0 + 0.71) * 256)); // 3-й столбец (88 px)
    }
}
