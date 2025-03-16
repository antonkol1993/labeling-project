package com.anton.labeling.service.large;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelCardCreator {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Card");

        // Устанавливаем ширину столбцов
        sheet.setColumnWidth(1, (int) (((124 - 5) / 7.0 + 0.71) * 256));
        sheet.setColumnWidth(2, (int) (((88 - 5) / 7.0 + 0.71) * 256));
        sheet.setColumnWidth(3, (int) (((88 - 5) / 7.0 + 0.71) * 256));

        // Создаем строки
        for (int i = 1; i <= 11; i++) {
            if (sheet.getRow(i) == null) {
                sheet.createRow(i);
            }
        }

        // Устанавливаем высоту строк
        sheet.getRow(1).setHeightInPoints(73.5f);
        sheet.getRow(2).setHeightInPoints(69.0f);
        sheet.getRow(3).setHeightInPoints(35.25f);

        // Объединенные ячейки
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 2, 3));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 2, 3));

        // Создание и применение стилей
        for (int row = 1; row <= 11; row++) {
            Row sheetRow = sheet.getRow(row);
            for (int col = 1; col <= 3; col++) {
                Cell cell = sheetRow.createCell(col);

                // Применяем стиль
                CellStyle cellStyle = ExcelCardStyleCreator.createBorderedCellStyle(workbook, row, col);
                cell.setCellStyle(cellStyle);
            }
        }

        // Работа с изображениями
        ExcelImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Mfix.jpg", 1, 1, 2, 3);
        ExcelImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Screw.jpg", 2, 1, 3, 3);

        // Сохранение файла
        try (FileOutputStream fileOut = new FileOutputStream("ExcelCard.xlsx")) {
            workbook.write(fileOut);
        }

        workbook.close();
        System.out.println("Файл создан: ExcelCard.xlsx");
    }
}

