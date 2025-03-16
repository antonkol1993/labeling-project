package com.anton.labeling.service.largeBox;

import org.apache.poi.ss.usermodel.*;
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

}
