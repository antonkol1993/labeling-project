package com.anton.labeling.service.largeBox;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CardStyle {


    // Метод для создания стиля ячейки
    public static CellStyle createBorderedCellStyle(XSSFWorkbook workbook, XSSFSheet sheet, int row, int col) {
        CellStyle cellStyle = workbook.createCellStyle();
        if (row <= 3) { // для строк 1-3
            cellStyle.setBorderTop(BorderStyle.MEDIUM);
            cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            cellStyle.setBorderLeft(BorderStyle.MEDIUM);
            cellStyle.setBorderRight(BorderStyle.MEDIUM);
        } else {
            if (col == 1) cellStyle.setBorderLeft(BorderStyle.MEDIUM);
            if (col == 3) cellStyle.setBorderRight(BorderStyle.MEDIUM);
            if (row < 10) cellStyle.setBorderBottom(BorderStyle.THIN);
            if (row == 10) cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        }
        return cellStyle;
    }

    // Метод для добавления объединенных ячеек с проверкой на дублирование
    public static void addMergedRegions(XSSFSheet sheet) {
        // Массив объединенных областей
        int[][] mergedRegions = {
                {1, 1, 1, 3}, // B2:D2
                {2, 2, 1, 3}, // B3:D3
                {3, 3, 1, 3}, // B4:D4
                {4, 4, 2, 3}, // C5:D5
                {5, 5, 2, 3}, // C6:D6
                {9, 9, 2, 3}  // C9:D9
        };

        // Проходим по всем диапазонам и добавляем их, если они еще не существуют
        for (int[] region : mergedRegions) {
            CellRangeAddress newRegion = new CellRangeAddress(region[0], region[1], region[2], region[3]);
            boolean regionExists = false;

            // Проверяем, есть ли уже такой объединенный регион
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress existingRegion = sheet.getMergedRegion(i);
                if (newRegion.equals(existingRegion)) {
                    regionExists = true;
                    break;
                }
            }

            // Если региона нет, добавляем его
            if (!regionExists) {
                sheet.addMergedRegion(newRegion);
            }
        }
    }

}
