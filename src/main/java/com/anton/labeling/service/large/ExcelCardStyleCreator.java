package com.anton.labeling.service.large;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCardStyleCreator {
    public static CellStyle createBorderedCellStyle(XSSFWorkbook workbook, int row, int col) {
        // Создаем стиль ячейки
        CellStyle cellStyle = workbook.createCellStyle();

        // Устанавливаем границы
        if (row <= 3) {
            // Границы для строк 1-3
            cellStyle.setBorderTop(BorderStyle.MEDIUM);
            cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            cellStyle.setBorderLeft(BorderStyle.MEDIUM);
            cellStyle.setBorderRight(BorderStyle.MEDIUM);
        } else {
            // Границы для строк 4-10
            if (col == 1) {
                cellStyle.setBorderLeft(BorderStyle.MEDIUM);
                cellStyle.setBorderRight(BorderStyle.THIN);
                if (row < 11) cellStyle.setBorderBottom(BorderStyle.THIN);
            } else if (col == 2) {
                if (row < 11) cellStyle.setBorderBottom(BorderStyle.THIN);
            } else if (col == 3) {
                cellStyle.setBorderRight(BorderStyle.MEDIUM);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                if (row < 11) cellStyle.setBorderBottom(BorderStyle.THIN);
            }
        }

        // Особые условия для последней строки (11)
        if (row == 11) {
            if (col == 1) {
                cellStyle.setBorderLeft(BorderStyle.MEDIUM);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            } else if (col == 2) {
                cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            } else if (col == 3) {
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.MEDIUM);
                cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            }
        }

        return cellStyle;
    }
}