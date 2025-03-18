package com.anton.labeling.service.irrelevantOld;


import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;


public class CardCreator {


    // Метод для создания и сохранения карточки
    public void createCard(XSSFWorkbook workbook, XSSFSheet sheet, ItemLargeBox item) throws IOException {
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
        CardStyle.addMergedRegions(sheet);

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
        celFiller.fillCells(workbook, sheet, item); // Добавляем вызов метода для заполнения

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





}