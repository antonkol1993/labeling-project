package com.anton.labeling.controller;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ExcelDataReader {

    private static int totalEmptyRows = 0;  // Общее количество пустых строк между блоками данных

    public static List<ItemLargeBox> readExcel(String filePath) throws IOException {
        totalEmptyRows = 0;  // Сбрасываем счетчик перед началом чтения

        List<ItemLargeBox> items = new ArrayList<>();
        FileInputStream file = new FileInputStream(new File(filePath));
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0); // Первый лист

        int emptyRowCounter = 0; // Количество пустых строк
        boolean foundDataBlock = false; // Флаг, чтобы начать подсчет пустых строк после первого блока данных

        boolean headersProcessed = false; // Флаг, обработаны ли заголовки

        for (Row row : sheet) {
            if (row.getRowNum() < 2) continue; // Пропускаем первые 2 строки (заголовки)

            if (!headersProcessed) {
                processHeaders(row); // Сохраняем заголовки
                headersProcessed = true;
                continue;
            }

            if (isRowEmpty(row)) {
                emptyRowCounter++; // Увеличиваем счетчик пустых строк
                continue; // Переходим к следующей строке
            }

            // Если были пустые строки, добавляем их к общему количеству
            if (emptyRowCounter > 0) {
                totalEmptyRows += emptyRowCounter;
                emptyRowCounter = 0; // Сбрасываем счетчик пустых строк после блока данных
            }

            // Обрабатываем блок данных
            processDataBlock(row, items);
            foundDataBlock = true; // После первого блока данных начинаем отслеживать пустые строки между блоками
        }

        workbook.close();
        file.close();
        return items;
    }

    public static Integer totalEmptyRows() {
        return totalEmptyRows;
    }

    // Метод для обработки заголовков
    private static void processHeaders(Row row) {
        System.out.println("Заголовки: ");
        for (Cell cell : row) {
            System.out.print(cell.getCellType() + " | ");
        }
        System.out.println("\n-------------------------");
    }

    // Метод для обработки блока данных
    private static void processDataBlock(Row row, List<ItemLargeBox> items) {
        ItemLargeBox item = new ItemLargeBox();
        item.setName(getCellValue(row.getCell(0)));
        item.setSize(getCellValue(row.getCell(1)));
        item.setQuantityInBox(getCellValue(row.getCell(2)));
        item.setMarking(getCellValue(row.getCell(3)));
        item.setOrder(getCellValue(row.getCell(4)));

        item.setNameAndSize(item.getName() + "\n" + item.getSize());

        items.add(item);
    }

    // Проверка, пустая ли строка
    private static boolean isRowEmpty(Row row) {
        if (row == null) return true;
        for (Cell cell : row) {
            if (cell != null && cell.getCellType() != CellType.BLANK && cell.getCellType() != CellType._NONE) {
                return false; // Строка не пуста, если хотя бы одна ячейка не пустая
            }
        }
        return true; // Строка пуста, если все ячейки пустые
    }

    // Безопасное получение значения ячейки
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString(); // Если дата, преобразуем в строку
                }
                return String.valueOf(cell.getNumericCellValue()); // Число преобразуем в строку
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}
