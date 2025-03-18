package com.anton.newpackage;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class DynamicExcelGenerator {
    private final Sheet sheet;
    private final int startRow;
    private final int startCol;

    public DynamicExcelGenerator(Sheet sheet, int startRow, int startCol) {
        this.sheet = sheet;
        this.startRow = startRow;
        this.startCol = startCol;
    }

    public void addCard() {
        // Определяем стили
        Workbook workbook = sheet.getWorkbook();
        CellStyle style1 = createCellStyle(workbook, "Calibri", false, IndexedColors.WHITE, BorderStyle.MEDIUM, HorizontalAlignment.CENTER);
        CellStyle style2 = createCellStyle(workbook, "Arial", true, IndexedColors.WHITE, BorderStyle.MEDIUM, HorizontalAlignment.CENTER);
        CellStyle style3 = createCellStyle(workbook, "Arial", true, IndexedColors.WHITE, BorderStyle.THIN, HorizontalAlignment.CENTER);
        CellStyle style4 = createCellStyle(workbook, "Arial", true, IndexedColors.WHITE, BorderStyle.THIN, HorizontalAlignment.GENERAL);

        // Устанавливаем ширину столбцов и высоту строк
        setColumnWidths(sheet, startCol);
        setRowHeights(sheet, startRow);

        // Заполняем таблицу относительно начальной точки
        createMergedCell(sheet, startRow, startCol, startRow, startCol + 2, "Label of organization", style1);
        createMergedCell(sheet, startRow + 1, startCol, startRow + 1, startCol + 2, "Photo of element", style2);
        createMergedCell(sheet, startRow + 2, startCol, startRow + 2, startCol + 2, "#REF!", style2);
        createCell(sheet, startRow + 3, startCol, "Marking", style4);
        createMergedCell(sheet, startRow + 3, startCol + 1, startRow + 3, startCol + 2, "#REF!", style3);
        createCell(sheet, startRow + 4, startCol, "РАЗМЕР/Size", style4);
        createMergedCell(sheet, startRow + 4, startCol + 1, startRow + 4, startCol + 2, "#REF!", style3);
        createCell(sheet, startRow + 5, startCol, "Кол-во/Q-ty", style4);
        createCell(sheet, startRow + 5, startCol + 1, "", style3);
        createCell(sheet, startRow + 5, startCol + 2, "Шт / PCS", style4);
        createCell(sheet, startRow + 6, startCol, "Кол-во в упак/шт.", style4);
        createCell(sheet, startRow + 6, startCol + 1, "#REF!", style3);
        createCell(sheet, startRow + 6, startCol + 2, "Шт / PCS", style4);
        createCell(sheet, startRow + 7, startCol, "Вес упак Кг/Kgs", style4);
        createCell(sheet, startRow + 7, startCol + 1, "", style3);
        createCell(sheet, startRow + 7, startCol + 2, "Кг/Kgs", style4);
        createCell(sheet, startRow + 8, startCol, "", style4);
        createMergedCell(sheet, startRow + 8, startCol + 1, startRow + 8, startCol + 2, "Сделано в КНР", style4);
        createCell(sheet, startRow + 9, startCol, "ORDER:", style4);
        createMergedCell(sheet, startRow + 9, startCol + 1, startRow + 9, startCol + 2, "#REF!", style4);
    }

    private static CellStyle createCellStyle(Workbook workbook, String fontName, boolean bold, IndexedColors bgColor, BorderStyle border, HorizontalAlignment alignment) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(alignment);
        style.setFillForegroundColor(bgColor.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderTop(border);
        style.setBorderBottom(border);
        style.setBorderLeft(border);
        style.setBorderRight(border);

        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontName(fontName);
        font.setBold(bold);
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);

        return style;
    }

    private static void createMergedCell(Sheet sheet, int rowStart, int colStart, int rowEnd, int colEnd, String text, CellStyle style) {
        sheet.addMergedRegion(new CellRangeAddress(rowStart - 1, rowEnd - 1, colStart - 1, colEnd - 1));

        for (int row = rowStart; row <= rowEnd; row++) {
            for (int col = colStart; col <= colEnd; col++) {
                createCell(sheet, row, col, "", style);
            }
        }

        createCell(sheet, rowStart, colStart, text, style);
    }

    private static void createCell(Sheet sheet, int row, int col, String text, CellStyle style) {
        Row sheetRow = sheet.getRow(row - 1);
        if (sheetRow == null) {
            sheetRow = sheet.createRow(row - 1);
        }
        Cell cell = sheetRow.createCell(col - 1);
        cell.setCellValue(text);
        cell.setCellStyle(style);
    }

    private static void setColumnWidths(Sheet sheet, int startCol) {
        sheet.setColumnWidth(startCol - 1, (int) (((124 - 5) / 7.0 + 0.71) * 256));
        sheet.setColumnWidth(startCol, (int) (((88 - 5) / 7.0 + 0.71) * 256));
        sheet.setColumnWidth(startCol + 1, (int) (((88 - 5) / 7.0 + 0.71) * 256));
    }

    private static void setRowHeights(Sheet sheet, int startRow) {
        setRowHeight(sheet, startRow, 73.5f);
        setRowHeight(sheet, startRow + 1, 69.0f);
        setRowHeight(sheet, startRow + 2, 35.25f);
    }

    private static void setRowHeight(Sheet sheet, int rowIndex, float height) {
        Row row = sheet.getRow(rowIndex - 1);
        if (row == null) {
            row = sheet.createRow(rowIndex - 1);
        }
        row.setHeightInPoints(height);
    }
}