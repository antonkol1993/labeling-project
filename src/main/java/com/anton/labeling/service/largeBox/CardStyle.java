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

    // Метод для добавления объединенных ячеек с проверкой на дублирование
    public static void addMergedRegions(XSSFSheet sheet, int startRow, int startCol) {
        int[][] mergedRegions = {
                {startRow, startRow, startCol, startCol + 2}, // 1 строка
                {startRow + 1, startRow + 1, startCol, startCol + 2}, // 2 строка
                {startRow + 2, startRow + 2, startCol, startCol + 2}, // 3 строка
                {startRow + 3, startRow + 3, startCol + 1, startCol + 2}, // 4 строка (правая часть)
                {startRow + 4, startRow + 4, startCol + 1, startCol + 2}, // 5 строка (C5:D5 - объединить!)

                {startRow + 8, startRow + 8, startCol + 1, startCol + 2},  // 9 строка
                {startRow + 9, startRow + 9, startCol + 1, startCol + 2}  // 10 строка
        };

        // Удаляем объединение 6 строки (C6:D6)
        sheet.getMergedRegions().removeIf(region ->
                region.getFirstRow() == startRow + 5 && region.getFirstColumn() == startCol + 1 &&
                        region.getLastRow() == startRow + 5 && region.getLastColumn() == startCol + 2
        );

        // Добавляем правильные объединенные области
        for (int[] region : mergedRegions) {
            CellRangeAddress newRegion = new CellRangeAddress(region[0], region[1], region[2], region[3]);
            boolean regionExists = sheet.getMergedRegions().stream().anyMatch(existing -> existing.equals(newRegion));

            if (!regionExists) {
                sheet.addMergedRegion(newRegion);
            }
        }
    }
}
