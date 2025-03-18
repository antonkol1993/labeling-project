package com.anton.newpackage;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelStyleReader {
    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream(new File("excel-example/Example.xlsx"));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0); // Получаем первый лист

        for (Row row : sheet) {
            for (Cell cell : row) {
                CellStyle style = cell.getCellStyle(); // Получаем стиль ячейки
                Font font = workbook.getFontAt(style.getFontIndexAsInt());

                int rowIndex = row.getRowNum();
                int colIndex = cell.getColumnIndex();

                System.out.println("Ячейка [" + rowIndex + "," + colIndex + "]: " + cell.toString());

                // Проверка объединения ячейки
                CellRangeAddress mergedRegion = getMergedRegion(sheet, rowIndex, colIndex);
                if (mergedRegion != null) {
                    System.out.println("  - Ячейка объединена с диапазоном: ("
                            + mergedRegion.getFirstRow() + "," + mergedRegion.getFirstColumn() + ") -> ("
                            + mergedRegion.getLastRow() + "," + mergedRegion.getLastColumn() + ")");
                } else {
                    System.out.println("  - Ячейка не объединена");
                }

                // Проверяем жирный ли текст
                if (font.getBold()) {
                    System.out.println("  - Шрифт: " + font.getFontName() + " (жирный)");
                } else {
                    System.out.println("  - Шрифт: " + font.getFontName());
                }

                // Выравнивание текста
                System.out.println("  - Выравнивание: " + style.getAlignment());

                // Цвет текста
                System.out.println("  - Цвет шрифта: " + font.getColor());

                // Границы (пример для верхней)
                System.out.println("  - Граница сверху: " + style.getBorderTop());

                // Фоновый цвет
                System.out.println("  - Фон ячейки (индекс): " + style.getFillForegroundColor());
            }
        }

        workbook.close();
        file.close();
    }

    private static CellRangeAddress getMergedRegion(Sheet sheet, int row, int col) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.isInRange(row, col)) {
                return mergedRegion;
            }
        }
        return null;
    }
}
