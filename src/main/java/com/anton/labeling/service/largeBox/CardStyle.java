package com.anton.labeling.service.largeBox;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CardStyle {

    public static CellStyle createBorderedCellStyle(XSSFWorkbook workbook, int row, int col, int startRow, int startCol) {
        CellStyle cellStyle = workbook.createCellStyle();
        int relativeRow = row - startRow;
        int relativeCol = col - startCol + 1; // Приводим к диапазону 1-3

        // Границы для заголовка карточки
        if (relativeRow <= 3) {
            cellStyle.setBorderTop(BorderStyle.MEDIUM);
            cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            cellStyle.setBorderLeft(BorderStyle.MEDIUM);
            cellStyle.setBorderRight(BorderStyle.MEDIUM);
        } else {
            if (relativeCol == 1) {
                cellStyle.setBorderLeft(BorderStyle.MEDIUM);
                if (relativeRow >= 4 && relativeRow <= 10) cellStyle.setBorderRight(BorderStyle.THIN);
            }
            if (relativeCol == 2) {
                if ((relativeRow >= 6 && relativeRow <= 8) || relativeRow == 10) {
                    cellStyle.setBorderRight(BorderStyle.THIN);
                }
            }
            if (relativeCol == 3) cellStyle.setBorderRight(BorderStyle.MEDIUM);
            if (relativeRow < 10) cellStyle.setBorderBottom(BorderStyle.THIN);
            if (relativeRow == 10) cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        }

        return cellStyle;
    }



    // Добавление объединенных ячеек (адаптирован для динамических карточек)
    public static void addMergedRegions(XSSFSheet sheet, int startRow, int startCol) {
        int[][] mergedRegions = {
                {1, 1, startCol, startCol + 2}, // Заголовок B2:D2, F2:H2...
                {2, 2, startCol, startCol + 2},
                {3, 3, startCol, startCol + 2},
                {4, 4, startCol + 1, startCol + 2},
                {5, 5, startCol + 1, startCol + 2},
                {9, 9, startCol + 1, startCol + 2}
        };

        for (int[] region : mergedRegions) {
            CellRangeAddress newRegion = new CellRangeAddress(region[0], region[1], region[2], region[3]);
            boolean regionExists = sheet.getMergedRegions().stream().anyMatch(existing -> existing.equals(newRegion));

            if (!regionExists) {
                sheet.addMergedRegion(newRegion);
            }
        }
    }
}
