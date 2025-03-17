package com.anton.newbot;


import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class CardCreator {
    private final XSSFWorkbook workbook;
    private final XSSFSheet sheet;

    public CardCreator(XSSFWorkbook workbook, XSSFSheet sheet) {
        this.workbook = workbook;
        this.sheet = sheet;
    }

    public void createCard(ItemLargeBox item, int startRow, int startCol) {
        CellFiller cellFiller = new CellFiller(workbook, sheet);

        // Установка ширины столбцов
        setColumnWidths(sheet, startCol);

        // Создание строк, если их нет
        for (int i = startRow; i < startRow + 10; i++) {
            if (sheet.getRow(i) == null) {
                sheet.createRow(i);
            }
        }

        // Установка высоты строк
        sheet.getRow(startRow).setHeightInPoints(73.5f);
        sheet.getRow(startRow + 1).setHeightInPoints(69.0f);
        sheet.getRow(startRow + 2).setHeightInPoints(35.25f);

        // Добавление объединённых областей
        CardStyle.addMergedRegions(sheet, startRow, startCol);

        // Заполнение данными
        cellFiller.fillCellsWithData(item, startRow, startCol);

        // Сохранение файла
        try (FileOutputStream fileOut = new FileOutputStream("output.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void setColumnWidths(XSSFSheet sheet, int startCol) {
        sheet.setColumnWidth(startCol, 4200);
        sheet.setColumnWidth(startCol + 1, 3000);
        sheet.setColumnWidth(startCol + 2, 3000);
    }
}

