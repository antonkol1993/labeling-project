package com.anton.newbot;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class CardStyle {
    public static void addMergedRegions(XSSFSheet sheet, int startRow, int startCol) {
        int[][] mergedRegions = {
                {startRow, startRow, startCol, startCol + 2},
                {startRow + 1, startRow + 1, startCol, startCol + 2},
                {startRow + 2, startRow + 2, startCol, startCol + 2},
                {startRow + 3, startRow + 3, startCol + 1, startCol + 2},
                {startRow + 4, startRow + 4, startCol + 1, startCol + 2},
                {startRow + 8, startRow + 8, startCol + 1, startCol + 2},
                {startRow + 9, startRow + 9, startCol + 1, startCol + 2}
        };

        for (int[] region : mergedRegions) {
            sheet.addMergedRegion(new CellRangeAddress(region[0], region[1], region[2], region[3]));
        }
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
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Arial");  // Устанавливаем шрифт Arial
        font.setBold(true);
        font.setFontHeightInPoints(fontSize);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    public static CellStyle createLeftBoldStyle(Workbook workbook, short fontSize) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Arial");  // Устанавливаем шрифт Arial
        font.setBold(true);
        font.setFontHeightInPoints(fontSize);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }
}
