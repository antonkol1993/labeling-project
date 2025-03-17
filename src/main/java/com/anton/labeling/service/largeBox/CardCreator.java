package com.anton.labeling.service.largeBox;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class CardCreator {
    public void createCard(XSSFWorkbook workbook, XSSFSheet sheet, ItemLargeBox item, int startRow, int startCol) throws IOException {
        CellFiller cellFiller = new CellFiller(workbook, sheet);

        // 1. Устанавливаем ширину столбцов
        setColumnWidths(sheet, startCol);

        // 2. Создаем строки, если их нет
        for (int i = startRow; i < startRow + 10; i++) {
            if (sheet.getRow(i) == null) {
                sheet.createRow(i);
            }
        }

        // 3. Устанавливаем высоту строк
        sheet.getRow(startRow).setHeightInPoints(73.5f);
        sheet.getRow(startRow + 1).setHeightInPoints(69.0f);
        sheet.getRow(startRow + 2).setHeightInPoints(35.25f);





        // 6. Применяем стили к ячейкам
        for (int row = startRow; row <= startRow + 9; row++) {
            Row sheetRow = sheet.getRow(row);
            if (sheetRow == null) {
                sheetRow = sheet.createRow(row);
            }
            for (int col = startCol; col <= startCol + 2; col++) {
                Cell cell = sheetRow.createCell(col);
                CellStyle cellStyle = CardStyle.createBorderedCellStyle(workbook, startRow, startCol, row, col);
                cell.setCellStyle(cellStyle);
            }
        }
//        celFiller.fillCells(item,startRow,startCol);
        cellFiller.fillCellsWithData(item, startRow, startCol);

        // 4. Добавляем объединенные области
        CardStyle.addMergedRegions(sheet, startRow, startCol);
        // 7. Добавляем изображения
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Mfix.jpg",
                startRow, startCol, startRow + 1, startCol + 2, 420000, 150000);
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Screw.jpg",
                startRow + 1, startCol, startRow + 2, startCol + 2, 400000, 130000);

        // ✅ 5. Повторно устанавливаем ширину столбцов (исправление!)
        setColumnWidths(sheet, startCol);
    }


    // Новый метод для установки ширины столбцов
    private void setColumnWidths(XSSFSheet sheet, int startCol) {
        sheet.setColumnWidth(startCol, (int) (((124 - 5) / 7.0 + 0.71) * 256)); // 1-й столбец (124 px)
        sheet.setColumnWidth(startCol + 1, (int) (((88 - 5) / 7.0 + 0.71) * 256)); // 2-й столбец (88 px)
        sheet.setColumnWidth(startCol + 2, (int) (((88 - 5) / 7.0 + 0.71) * 256)); // 3-й столбец (88 px)
    }
}


