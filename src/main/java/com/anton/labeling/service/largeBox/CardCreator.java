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

        // Динамически устанавливаем ширину столбцов
        for (int i = 0; i < 3; i++) {
            sheet.setColumnWidth(startCol + i, (int) (((88 - 5) / 7.0 + 0.71) * 256));
        }

        // Динамически создаем строки, если они не существуют
        for (int i = startRow; i < startRow + 10; i++) {
            if (sheet.getRow(i) == null) {
                sheet.createRow(i);
            }
        }

        // Устанавливаем высоту строк
        sheet.getRow(startRow).setHeightInPoints(73.5f);
        sheet.getRow(startRow + 1).setHeightInPoints(69.0f);
        sheet.getRow(startRow + 2).setHeightInPoints(35.25f);

        // Добавляем объединенные области
        CardStyle.addMergedRegions(sheet, startRow, startCol);

        // Применяем стили к ячейкам
        for (int row = startRow; row <= startRow + 9; row++) {  // 10 строк
            Row sheetRow = sheet.getRow(row);
            if (sheetRow == null) {
                sheetRow = sheet.createRow(row);
            }
            for (int col = startCol; col <= startCol + 2; col++) {
                Cell cell = sheetRow.createCell(col);

                // Применяем стиль
                CellStyle cellStyle = CardStyle.createBorderedCellStyle(workbook, startRow, startCol, row, col);
                cell.setCellStyle(cellStyle);
            }
        }

// Добавляем объединенные области
        CardStyle.addMergedRegions(sheet, startRow, startCol);

        // Заполняем карточку данными
//        celFiller.fillCells(item, startRow, startCol);

        // Добавляем изображения
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Mfix.jpg",
                startRow, startCol, startRow + 1, startCol + 2, 420000, 150000);
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Screw.jpg",
                startRow + 1, startCol, startRow + 2, startCol + 2, 400000, 130000);
    }

}
