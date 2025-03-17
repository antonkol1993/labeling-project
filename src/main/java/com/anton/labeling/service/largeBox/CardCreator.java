package com.anton.labeling.service.largeBox;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class CardCreator {
    public void createCard(XSSFWorkbook workbook, XSSFSheet sheet, ItemLargeBox item, int startRow, int startCol) throws IOException {
        CellFiller cellFiller = new CellFiller(workbook, sheet);

        // Устанавливаем ширину столбцов
        CardStyle.setColumnWidths(sheet, startCol);

        // Создаем строки и задаем высоту
        for (int i = startRow; i < startRow + 10; i++) {
            if (sheet.getRow(i) == null) {
                sheet.createRow(i);
            }
        }
        sheet.getRow(startRow).setHeightInPoints(73.5f);
        sheet.getRow(startRow + 1).setHeightInPoints(69.0f);
        sheet.getRow(startRow + 2).setHeightInPoints(35.25f);

        // Применяем стили
        for (int row = startRow; row <= startRow + 9; row++) {
            for (int col = startCol; col <= startCol + 2; col++) {
                sheet.getRow(row).createCell(col)
                        .setCellStyle(CardStyle.createBorderedCellStyle(workbook, startRow, startCol, row, col));
            }
        }
        cellFiller.fillCellsWithData(item,startRow,startCol);

        // Объединяем ячейки
        CardStyle.addMergedRegions(sheet, startRow, startCol);

        // Добавляем изображения
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Mfix.jpg",
                startRow, startCol, startRow + 1, startCol + 2, 420000, 150000);
        ImageHandler.addImageToSheet(workbook, sheet, "src/main/resources/static/images/Screw.jpg",
                startRow + 1, startCol, startRow + 2, startCol + 2, 400000, 130000);
    }
}
