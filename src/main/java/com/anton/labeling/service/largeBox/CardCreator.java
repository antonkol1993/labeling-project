package com.anton.labeling.service.largeBox;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;


public class CardCreator {

    // Метод для создания и сохранения карточки
    public void createCard(XSSFWorkbook workbook, XSSFSheet sheet) throws IOException {
        // Устанавливаем ширину столбцов
        sheet.setColumnWidth(1, (int) (((124 - 5) / 7.0 + 0.71) * 256));
        sheet.setColumnWidth(2, (int) (((88 - 5) / 7.0 + 0.71) * 256));
        sheet.setColumnWidth(3, (int) (((88 - 5) / 7.0 + 0.71) * 256));

        // Создаем строки, если они еще не существуют
        for (int i = 1; i <= 10; i++) {
            if (sheet.getRow(i) == null) {
                sheet.createRow(i);
            }
        }

        // Устанавливаем высоту строк
        sheet.getRow(1).setHeightInPoints(73.5f);
        sheet.getRow(2).setHeightInPoints(69.0f);
        sheet.getRow(3).setHeightInPoints(35.25f);

        // Добавляем объединенные области
        addMergedRegions(sheet);

        // Применяем стиль ячеек
        for (int row = 1; row <= 10; row++) {  // Строки от 1 до 10
            Row sheetRow = sheet.getRow(row);
            for (int col = 1; col <= 3; col++) {
                Cell cell = sheetRow.createCell(col);

                // Применяем стиль
                CellStyle cellStyle = CardStyle.createBorderedCellStyle(workbook, sheet, row, col);
                cell.setCellStyle(cellStyle);
            }
        }

        // Заполнение ячеек в соответствии с заданными стилями
        fillCells(workbook, sheet); // Добавляем вызов метода для заполнения

        // Работа с изображениями
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Mfix.jpg",
                1, 1, 2, 3, 420000, 150000);
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Screw.jpg",
                2, 1, 3, 3, 400000, 130000);

        // Сохраняем файл
        try (FileOutputStream fileOut = new FileOutputStream("ExcelCard.xlsx")) {
            workbook.write(fileOut);
        }
        workbook.close();
        System.out.println("Excel создана: ExcelCard.xlsx");
    }

    // Метод для добавления объединенных ячеек с проверкой на дублирование
    private void addMergedRegions(XSSFSheet sheet) {
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

    private static void fillCells(XSSFWorkbook workbook, XSSFSheet sheet) {
        // 🔹 Жирный шрифт для строки 3 (B3)
        Font boldFontMain3Row = workbook.createFont();
        boldFontMain3Row.setBold(true);
        boldFontMain3Row.setFontHeightInPoints((short) 11);

        // 🔹 Заполнение B3 (жирный центр)
        Row row3 = sheet.getRow(3);
        if (row3 == null) row3 = sheet.createRow(3);
        Cell cell3 = row3.getCell(1);
        if (cell3 == null) cell3 = row3.createCell(1);
        cell3.setCellValue("Жирный Центр Item");

        CellStyle boldCenterStyle = CardStyle.createBorderedCellStyle(workbook, sheet, 3, 1);
        boldCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        boldCenterStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldCenterStyle.setFont(boldFontMain3Row);
        cell3.setCellStyle(boldCenterStyle);

        // 🔹 Жирный шрифт для левого выравнивания (B4-B10, кроме B9)
        Font boldFontDownRows = workbook.createFont();
        boldFontDownRows.setBold(true);
        boldFontDownRows.setFontHeightInPoints((short) 10);

        for (int rowNum = 4; rowNum <= 10; rowNum++) {
            if (rowNum == 9) continue;

            Row row = sheet.getRow(rowNum);
            if (row == null) row = sheet.createRow(rowNum);

            Cell cellB = row.getCell(1);
            if (cellB == null) cellB = row.createCell(1);
            cellB.setCellValue("Жирный лево Item");

            CellStyle boldLeftStyle = CardStyle.createBorderedCellStyle(workbook, sheet, rowNum, 1);
            boldLeftStyle.setAlignment(HorizontalAlignment.LEFT);
            boldLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            boldLeftStyle.setFont(boldFontDownRows);
            cellB.setCellStyle(boldLeftStyle);
        }

        // 🔹 Жирный шрифт для объединенных ячеек C5:D5 и C7
        Font boldFontCenter = workbook.createFont();
        boldFontCenter.setBold(true);
        boldFontCenter.setFontHeightInPoints((short) 10);

        CellStyle boldCenterMergedStyle = CardStyle.createBorderedCellStyle(workbook, sheet, 5, 2);
        boldCenterMergedStyle.setAlignment(HorizontalAlignment.CENTER);
        boldCenterMergedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldCenterMergedStyle.setFont(boldFontCenter);

        // 🔹 Заполнение C5:D5 (ячейки объединены)
        Row row5 = sheet.getRow(5);
        if (row5 == null) row5 = sheet.createRow(5);
        Cell cell5C = row5.getCell(2);
        if (cell5C == null) cell5C = row5.createCell(2);
        cell5C.setCellValue("Жирный Центр Item");
        cell5C.setCellStyle(boldCenterMergedStyle);

        // ✅ Проверяем, не объединены ли C5:D5 перед объединением
        if (!isMergedRegion(sheet, 5, 2, 5, 3)) {
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 2, 3));
        }

        // 🔹 Заполнение C7 (центр)
        Row row7 = sheet.getRow(7);
        if (row7 == null) row7 = sheet.createRow(7);
        Cell cell7C = row7.getCell(2);
        if (cell7C == null) cell7C = row7.createCell(2);
        cell7C.setCellValue("1 000");
        cell7C.setCellStyle(boldCenterMergedStyle);

        // 🔹 Жирный шрифт для текста "final left"
        Font boldFontFinalLeft = workbook.createFont();
        boldFontFinalLeft.setBold(true);
        boldFontFinalLeft.setFontHeightInPoints((short) 10);

        CellStyle boldLeftFinalStyle = workbook.createCellStyle();
        boldLeftFinalStyle.setAlignment(HorizontalAlignment.LEFT);
        boldLeftFinalStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldLeftFinalStyle.setFont(boldFontFinalLeft);

        // 🔹 Заполнение D6-D8 (текст слева, жирный)
        for (int rowNum = 6; rowNum <= 8; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) row = sheet.createRow(rowNum);

            Cell cellD = row.getCell(3);
            if (cellD == null) cellD = row.createCell(3);
            cellD.setCellValue("final left");

            CellStyle borderedLeftStyle = CardStyle.createBorderedCellStyle(workbook, sheet, rowNum, 3);
            borderedLeftStyle.setAlignment(HorizontalAlignment.LEFT);
            borderedLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            borderedLeftStyle.setFont(boldFontFinalLeft);
            cellD.setCellStyle(borderedLeftStyle);
        }

        // 🔹 Заполнение C9-C10 (текст слева, жирный)
        for (int rowNum = 9; rowNum <= 10; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) row = sheet.createRow(rowNum);

            Cell cellC = row.getCell(2);
            if (cellC == null) cellC = row.createCell(2);
            cellC.setCellValue("final left");

            CellStyle borderedLeftStyle = CardStyle.createBorderedCellStyle(workbook, sheet, rowNum, 2);
            borderedLeftStyle.setAlignment(HorizontalAlignment.LEFT);
            borderedLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            borderedLeftStyle.setFont(boldFontFinalLeft);
            cellC.setCellStyle(borderedLeftStyle);
        }

    }

    /**
     * Проверяет, существует ли уже объединенный регион в заданном диапазоне.
     */
    private static boolean isMergedRegion(XSSFSheet sheet, int firstRow, int firstCol, int lastRow, int lastCol) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstRow() == firstRow && range.getFirstColumn() == firstCol &&
                    range.getLastRow() == lastRow && range.getLastColumn() == lastCol) {
                return true;
            }
        }
        return false;
    }


}