package com.anton.labeling.service.largeBox;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class CardCreator {

    // Метод для создания и сохранения карточки
    public void createCard(XSSFWorkbook workbook, XSSFSheet sheet, ItemLargeBox item, int startRow, int startCol) throws IOException {
        CelFiller celFiller = new CelFiller(workbook, sheet);

        // Устанавливаем ширину столбцов (с учетом startCol)
        for (int i = 0; i < 3; i++) {
            sheet.setColumnWidth(startCol + i, (int) (((88 - 5) / 7.0 + 0.71) * 256));
        }

        // Создаем строки, если они еще не существуют
        for (int i = startRow; i < startRow + 11; i++) {
            if (sheet.getRow(i) == null) {
                sheet.createRow(i);
            }
        }

        // Устанавливаем высоту строк (с учетом смещения startRow)
        sheet.getRow(startRow + 0).setHeightInPoints(73.5f);
        sheet.getRow(startRow + 1).setHeightInPoints(69.0f);
        sheet.getRow(startRow + 2).setHeightInPoints(35.25f);

        // Добавляем объединенные области с учетом смещения
        CardStyle.addMergedRegions(sheet, startRow, startCol);

        // Применяем стили к ячейкам
        for (int row = startRow; row < startRow + 11; row++) {
            Row sheetRow = sheet.getRow(row);
            for (int col = startCol; col < startCol + 3; col++) {
                Cell cell = sheetRow.createCell(col);

                // Применяем стиль (передаем startRow и startCol)
                CellStyle cellStyle = CardStyle.createBorderedCellStyle(workbook, row, col, startRow, startCol);
                cell.setCellStyle(cellStyle);
            }
        }

        // Заполняем карточку данными
        celFiller.fillCells(item, startRow, startCol);

        // Добавляем изображения (адаптировано под startRow/startCol)
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Mfix.jpg",
                startRow + 0, startCol, startRow + 1, startCol + 2, 420000, 150000);
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Screw.jpg",
                startRow + 1, startCol, startRow + 2, startCol + 2, 400000, 130000);

        // Сохраняем файл
        try (FileOutputStream fileOut = new FileOutputStream("ExcelCard.xlsx")) {
            workbook.write(fileOut);
        }
//        workbook.close();
//        System.out.println("Excel создана: ExcelCard.xlsx");
    }
}
