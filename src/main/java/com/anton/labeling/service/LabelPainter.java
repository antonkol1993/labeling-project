package com.anton.labeling.service;

import com.anton.labeling.objects.ParamsOfElement;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.ss.util.CellRangeAddress;


public class LabelPainter {

    public static void createLabel(ParamsOfElement params, String outputFilePath) throws IOException {
        // Создаём новый Excel-документ
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Label");

        // 1. Объединяем ячейки: строка 2 (индекс 1), столбцы B (1), C (2), D (3)
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 3));

        // 2. Создаём строку 2 (индекс 1)
        Row orgLabelRow = sheet.createRow(1);

        // Создаём ячейку B2 (индекс 1)
        Cell orgLabelCell = orgLabelRow.createCell(1);
        orgLabelCell.setCellValue("Label of organization"); // Устанавливаем текст

        // Стиль для центрирования текста
        CellStyle centerStyle = workbook.createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER); // Центрирование текста

        // 3. Стиль для толстой внешней границы (вокруг всех сторон)
        CellStyle borderStyle = workbook.createCellStyle();
        borderStyle.setBorderTop(BorderStyle.THICK);    // Верхняя граница
        borderStyle.setBorderBottom(BorderStyle.THICK); // Нижняя граница
        borderStyle.setBorderLeft(BorderStyle.THICK);   // Левая граница
        borderStyle.setBorderRight(BorderStyle.THICK);  // Правая граница
        borderStyle.setAlignment(HorizontalAlignment.CENTER); // Центрирование текста

        // Применяем стиль с границами и центрированием
        orgLabelCell.setCellStyle(borderStyle);

        // Остальная часть кода (как в предыдущем примере)
        int rowIndex = 2;

        // Фото элемента
        Row photoRow = sheet.createRow(rowIndex++);
        photoRow.createCell(1).setCellValue("Photo of element");

        // Пустая строка
        sheet.createRow(rowIndex++);

        // Маркировка
        Row markingRow = sheet.createRow(rowIndex++);
        markingRow.createCell(1).setCellValue("Marking");
        markingRow.createCell(2).setCellValue(params.getMarkingOfElement());

        // Размер
        Row sizeRow = sheet.createRow(rowIndex++);
        sizeRow.createCell(1).setCellValue("РАЗМЕР/Size");
        sizeRow.createCell(2).setCellValue(params.getSizeOfElement());

        // Количество
        Row quantityRow = sheet.createRow(rowIndex++);
        quantityRow.createCell(1).setCellValue("Кол-во/Q-ty");
        quantityRow.createCell(3).setCellValue("Шт / PCS");

        // Количество в упаковке
        Row quantityInBoxRow = sheet.createRow(rowIndex++);
        quantityInBoxRow.createCell(1).setCellValue("Кол-во в упак/шт.");
        quantityInBoxRow.createCell(2).setCellValue(params.getQuantityIntoBoxOfElement());
        quantityInBoxRow.createCell(3).setCellValue("Шт / PCS");

        // Вес упаковки
        Row weightRow = sheet.createRow(rowIndex++);
        weightRow.createCell(1).setCellValue("Вес упак Кг/Kgs");
        weightRow.createCell(3).setCellValue("Кг/Kgs");

        // Происхождение
        Row originRow = sheet.createRow(rowIndex++);
        originRow.createCell(2).setCellValue("Сделано в " + params.getOriginOfElement());

        // Заказ
        Row orderRow = sheet.createRow(rowIndex++);
        orderRow.createCell(1).setCellValue("ORDER:");
        orderRow.createCell(2).setCellValue(params.getOriginOfElement());

        // Автоматическое изменение ширины столбцов
        for (int i = 0; i < 4; i++) {
            sheet.autoSizeColumn(i);
        }

        // Сохраняем файл
        try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
            workbook.write(fileOut);
        }

        // Закрываем workbook
        workbook.close();
    }

    public static void main(String[] args) throws IOException {
        // Пример использования
        ParamsOfElement params = new ParamsOfElement();
        params.setMarkingOfElement("PZ");
        params.setSizeOfElement("4.2x13");
        params.setQuantityIntoBoxOfElement("1000");
        params.setOriginOfElement("24HS10214P");
        params.setOriginOfElement("КНР");

        createLabel(params, "output_label.xlsx");
    }
}
