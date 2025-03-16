package com.anton.labeling.service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;


public class ExcelCardCreatorOLDDDDDD {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Card");

        // Устанавливаем ширину столбцов
        int columnWidth1 = 124;
        int columnWidth2 = 88;
        int columnWidth3 = 88;
        sheet.setColumnWidth(1, (int) (((columnWidth1 - 5) / 7.0 + 0.71) * 256)); // B
        sheet.setColumnWidth(2, (int) (((columnWidth2 - 5) / 7.0 + 0.71) * 256)); // C
        sheet.setColumnWidth(3, (int) (((columnWidth3 - 5) / 7.0 + 0.71) * 256)); // D

        // Создаем строки перед установкой высоты
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
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 3)); // B2:D2
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 3)); // B3:D3
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 1, 3)); // B4:D4
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 2, 3)); // C5:D5
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 2, 3)); // C6:D6

        for (int row = 1; row <= 11; row++) {
            Row sheetRow = sheet.getRow(row);
            for (int col = 1; col <= 3; col++) {
                Cell cell = sheetRow.createCell(col);

                // Создаём стиль для текущей ячейки
                CellStyle cellStyle = workbook.createCellStyle();

                if (row <= 3) {
                    // С 1 по 3 строку - везде граница MEDIUM
                    cellStyle.setBorderTop(BorderStyle.MEDIUM);
                    cellStyle.setBorderBottom(BorderStyle.MEDIUM);
                    cellStyle.setBorderLeft(BorderStyle.MEDIUM);
                    cellStyle.setBorderRight(BorderStyle.MEDIUM);
                } else {
                    // Настройки границ для строк 4–10
                    if (col == 1) { // Column 1
                        cellStyle.setBorderLeft(BorderStyle.MEDIUM);  // Толстая слева
                        cellStyle.setBorderRight(BorderStyle.THIN);   // Тонкая справа
                        if (row < 11) cellStyle.setBorderBottom(BorderStyle.THIN); // Тонкая снизу
                    } else if (col == 2) { // Column 2
                        if (row < 11) cellStyle.setBorderBottom(BorderStyle.THIN); // Тонкая снизу
                    } else if (col == 3) { // Column 3
                        cellStyle.setBorderRight(BorderStyle.MEDIUM);  // Толстая справа
                        cellStyle.setBorderLeft(BorderStyle.THIN);     // Тонкая слева
                        if (row < 11) cellStyle.setBorderBottom(BorderStyle.THIN); // Тонкая снизу
                    }

                    // Особые условия для последней строки (11)
                    if (row == 11) {
                        if (col == 1) {
                            cellStyle.setBorderLeft(BorderStyle.MEDIUM);   // Толстая слева
                            cellStyle.setBorderRight(BorderStyle.THIN);    // Тонкая справа
                            cellStyle.setBorderBottom(BorderStyle.MEDIUM); // Толстая снизу
                        } else if (col == 2) {
                            cellStyle.setBorderBottom(BorderStyle.MEDIUM); // Толстая снизу
                        } else if (col == 3) {
                            cellStyle.setBorderLeft(BorderStyle.THIN);     // Тонкая слева
                            cellStyle.setBorderRight(BorderStyle.MEDIUM);  // Толстая справа
                            cellStyle.setBorderBottom(BorderStyle.MEDIUM); // Толстая снизу
                        }
                    }
                }

                cell.setCellStyle(cellStyle);
            }
        }


        // работа с ихображением

        // Загружаем изображение MFIX
        InputStream inputStream = new FileInputStream("src/main/resources/static/images/Mfix.jpg");
        byte[] imageBytes = IOUtils.toByteArray(inputStream);
        int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
        inputStream.close();

        // Создаём объект для работы с изображением
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = new XSSFClientAnchor();

        // Настраиваем положение изображения
        anchor.setCol1(1); // Колонка B
        anchor.setRow1(1); // Строка 2
        anchor.setCol2(3); // До колонки D
        anchor.setRow2(2); // До строки 3

        anchor.setDx1(400000); // Смещение вправо
        anchor.setDy1(200000);  // Смещение вниз


        // Вставляем изображение
        Picture pictureMFix = drawing.createPicture(anchor, pictureIdx);
        pictureMFix.resize(); // Автоматический размер

        // Загружаем изображение Screw
        inputStream = new FileInputStream("src/main/resources/static/images/Screw.jpg");
        imageBytes = IOUtils.toByteArray(inputStream);
        pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
        inputStream.close();

        // Создаём объект для работы с изображением
         drawing = sheet.createDrawingPatriarch();
         anchor = new XSSFClientAnchor();

        // Настраиваем положение изображения
        anchor.setCol1(1); // Колонка B
        anchor.setRow1(2); // Строка 2
        anchor.setCol2(3); // До колонки D
        anchor.setRow2(3); // До строки 3

        anchor.setDx1(420000); // Смещение вправо
        anchor.setDy1(130000);  // Смещение вниз


        // Вставляем изображение
        pictureMFix = drawing.createPicture(anchor, pictureIdx);
        pictureMFix.resize(); // Автоматический размер



        // Сохранение файла
            try (FileOutputStream fileOut = new FileOutputStream("ExcelCard.xlsx")) {
                workbook.write(fileOut);
            }
            workbook.close();
            System.out.println("Файл создан: ExcelCard.xlsx");


    }
}
