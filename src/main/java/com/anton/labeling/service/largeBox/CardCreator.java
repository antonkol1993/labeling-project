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
        CelFiller celFiller = new CelFiller(workbook, sheet);
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

    private void fillCells(XSSFWorkbook workbook, XSSFSheet sheet) {
        CelFiller cellFiller = new CelFiller(workbook, sheet);

        // 🔹 Жирный шрифт для заголовков
        Font arialFontMain = cellFiller.createArialFont((short) 11);

        // Заполнение B3 (жирный центр)
        cellFiller.setCellValueWithStyle(3, 1, "Жирный Центр Item", arialFontMain, HorizontalAlignment.CENTER);

        // 🔹 Жирный шрифт для B4-B10 (кроме B9) с левым выравниванием
        Font arialFontDownRows = cellFiller.createArialFont((short) 10);
        for (int rowNum = 4; rowNum <= 10; rowNum++) {
            if (rowNum == 9) continue; // Пропускаем строку 9
            cellFiller.setCellValueWithStyle(rowNum, 1, "Жирный лево Item", arialFontDownRows, HorizontalAlignment.LEFT);
        }

        // 🔹 Жирный шрифт для C5:D5, C7 (центр)
        Font arialFontCenter = cellFiller.createArialFont((short) 10);
        cellFiller.setCellValueWithMergedStyle(5, 2, "Жирный Центр Item", arialFontCenter, HorizontalAlignment.CENTER, 5, 3);
        cellFiller.setCellValueWithStyle(7, 2, "Жирный Центр Item", arialFontCenter, HorizontalAlignment.CENTER);

        // 🔹 Заполнение D6-D8, C9-C10 (текст слева, жирный)
        Font arialFontFinalLeft = cellFiller.createArialFont((short) 10);
        for (int rowNum = 6; rowNum <= 8; rowNum++) {
            cellFiller.setCellValueWithStyle(rowNum, 3, "final left", arialFontFinalLeft, HorizontalAlignment.LEFT);
        }
        for (int rowNum = 9; rowNum <= 10; rowNum++) {
            cellFiller.setCellValueWithStyle(rowNum, 2, "final left", arialFontFinalLeft, HorizontalAlignment.LEFT);
        }
    }


}