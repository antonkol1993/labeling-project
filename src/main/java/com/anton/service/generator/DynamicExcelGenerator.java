package com.anton.service.generator;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.List;

public class DynamicExcelGenerator {

    private final Workbook workbook;
    private final Sheet sheet;
    private final boolean isXSSF;
    private int startRow = 2;
    private int startCol = 2;

    public DynamicExcelGenerator(Workbook workbook, Sheet sheet) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.isXSSF = workbook instanceof XSSFWorkbook;
    }

    public void generateCardsFromBlocks(List<List<ItemLargeBox>> dataBlocks) throws IOException {
        int tempCol = startCol;
        for (List<ItemLargeBox> block : dataBlocks) {
            for (ItemLargeBox item : block) {
                addCard(item);
                startCol += 4; // Сдвигаем вправо на 4 колонки
            }
            startRow += 12; // Сдвигаем вниз на 12 строк
            startCol = tempCol; // Возвращаем колонку в начало
        }
    }

    public void addCard(ItemLargeBox itemLargeBox) throws IOException {
        CellStyle style1 = createCellStyle("Arial", false, BorderStyle.MEDIUM, HorizontalAlignment.CENTER, (short) 10);
        CellStyle style2 = createCellStyle("Arial", true, BorderStyle.MEDIUM, HorizontalAlignment.CENTER, (short) 11);
        CellStyle style3 = createCellStyle("Arial", true, BorderStyle.THIN, HorizontalAlignment.CENTER, (short) 10);
        CellStyle style4 = createCellStyle("Arial", true, BorderStyle.THIN, HorizontalAlignment.GENERAL, (short) 10);

        setColumnWidths(sheet, startCol); // Корректный сдвиг вправо
        setRowHeights(sheet, startRow); // Корректный сдвиг вниз

        createMergedCell(startRow, startCol + 1, startRow, startCol + 3, "", style1);
        createMergedCell(startRow + 1, startCol + 1, startRow + 1, startCol + 3, "", style2);
        createMergedCell(startRow + 2, startCol + 1, startRow + 2, startCol + 3, itemLargeBox.getNameAndSize(), style2);

        createCell(startRow + 3, startCol + 1, "Marking", style4);
        createMergedCell(startRow + 3, startCol + 2, startRow + 3, startCol + 3, itemLargeBox.getMarking(), style3);

        createCell(startRow + 4, startCol + 1, "РАЗМЕР/Size", style4);
        createMergedCell(startRow + 4, startCol + 2, startRow + 4, startCol + 3, itemLargeBox.getSize(), style3);

        createCell(startRow + 5, startCol + 1, "", style4);
        createMergedCell(startRow + 5, startCol + 2, startRow + 5, startCol + 3, "", style3);

        createCell(startRow + 6, startCol + 1, "Кол-во в упак/шт.", style4);
        createCell(startRow + 6, startCol + 2, itemLargeBox.getQuantityInBox(), style3);
        createCell(startRow + 6, startCol + 3, "Шт / PCS", style4);

        createCell(startRow + 7, startCol + 1, "Вес упак Кг/Kgs", style4);
        createCell(startRow + 7, startCol + 2, "", style3);
        createCell(startRow + 7, startCol + 3, "Кг/Kgs", style4);

        createCell(startRow + 8, startCol + 1, "", style4);
        createMergedCell(startRow + 8, startCol + 2, startRow + 8, startCol + 3, "Сделано в КНР", style4);

        createCell(startRow + 9, startCol + 1, "ORDER:", style4);
        createMergedCell(startRow + 9, startCol + 2, startRow + 9, startCol + 3, itemLargeBox.getOrder(), style4);

        // 🔹 Авторазмер всех строк карточки
        for (int i = startRow + 2; i <= startRow + 9; i++) {
            autoSizeRow(sheet, i);
        }


        // Добавление изображений
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Mfix.jpg",
                startRow - 1, startCol, startRow - 1, startCol + 2, 420000, 150000);
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Example.jpg",
                startRow, startCol, startRow, startCol + 2, 400000, 130000);
    }


    private CellStyle createCellStyle(String fontName, boolean bold, BorderStyle border, HorizontalAlignment alignment, short fontSize) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setBold(bold);
        font.setFontHeightInPoints(fontSize);
        style.setFont(font);
        style.setAlignment(alignment);
        style.setBorderBottom(border);
        style.setBorderTop(border);
        style.setBorderLeft(border);
        style.setBorderRight(border);
        style.setWrapText(true); // Включаем перенос текста
        return style;
    }

    private void setColumnWidths(Sheet sheet, int startCol) {
        sheet.setColumnWidth(startCol, (int) (((124 - 5) / 7.0 + 0.71) * 256));
        sheet.setColumnWidth(startCol + 1, (int) (((88 - 5) / 7.0 + 0.71) * 256));
        sheet.setColumnWidth(startCol + 2, (int) (((88 - 5) / 7.0 + 0.71) * 256));
    }

    private void setRowHeights(Sheet sheet, int startRow) {
        setRowHeight(sheet, startRow, 73.5f);
        setRowHeight(sheet, startRow + 1, 69.0f);
        setRowHeight(sheet, startRow + 2, 35.25f);
    }

    private void setRowHeight(Sheet sheet, int rowIndex, float height) {
        Row row = sheet.getRow(rowIndex - 1);
        if (row == null) {
            row = sheet.createRow(rowIndex - 1);
        }
        row.setHeightInPoints(height);
    }

    private void createCell(int row, int col, String value, CellStyle style) {
        Row sheetRow = sheet.getRow(row - 1);
        if (sheetRow == null) {
            sheetRow = sheet.createRow(row - 1);
        }
        Cell cell = sheetRow.createCell(col - 1);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private void createMergedCell(int startRow, int startCol, int endRow, int endCol, String value, CellStyle style) {
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, endRow - 1, startCol - 1, endCol - 1));

        for (int row = startRow; row <= endRow; row++) {
            for (int col = startCol; col <= endCol; col++) {
                createCell(row, col, "", style);
            }
        }

        createCell(startRow, startCol, value, style);
    }

    private void autoSizeRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex - 1);
        if (row != null) {
            int maxTextLength = 0;

            // Проверяем все колонки в строке
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    int textLength = cell.getStringCellValue().length();
                    maxTextLength = Math.max(maxTextLength, textLength);
                }
            }

            int lineCount = (int) Math.ceil(maxTextLength / 20.0); // 20 символов в строке
            row.setHeightInPoints(lineCount * sheet.getDefaultRowHeightInPoints()); // Авторазмер
        }
    }


}
