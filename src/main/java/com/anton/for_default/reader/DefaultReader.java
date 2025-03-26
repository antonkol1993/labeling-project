package com.anton.for_default.reader;

import com.anton.for_default.obj.DefaultItem;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DefaultReader {

    public static void main(String[] args) {
        String filePath = "excel-example/DataFromInvoice for example .xlsx"; // Укажите путь к файлу

        DefaultReader reader = new DefaultReader();
        try {
            List<List<DefaultItem>> dataBlocks = reader.readExcel(filePath);

            // Логируем результат
            for (int i = 0; i < dataBlocks.size(); i++) {
                System.out.printf("Блок данных #%d%n", i + 1);
                System.out.println();
                for (DefaultItem item : dataBlocks.get(i)) {
                    System.out.println(item.toString());
                }
                System.out.println("----------------------");
            }

        } catch (IOException e) {
            logger.error("Ошибка при чтении файла: {}", e.getMessage());
        }
    }


    private final List<List<DefaultItem>> dataBlocks = new ArrayList<>();
    private List<DefaultItem> currentBlock = new ArrayList<>();

    private static final Logger logger = LoggerFactory.getLogger(DefaultReader.class);

    public List<List<DefaultItem>> readExcel(String filePath) throws IOException {
        FileInputStream file = new FileInputStream(new File(filePath));
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);

        // Обрабатываем все строки в листе
        for (Row row : sheet) {
            logger.debug("Читаем строку #{}", row.getRowNum() + 1);  // Логируем номер строки
            if (row.getRowNum() < 2) continue; // Пропускаем заголовки (читаем с 3 строки)

            // Обрабатываем строку
            DefaultItem item = processDataBlock(row);

            if (item != null) {
                logger.info("Обработан элемент: {}", item); // Логируем обработанный элемент
                currentBlock.add(item);
            } else {
                if (!currentBlock.isEmpty()) {
                    dataBlocks.add(new ArrayList<>(currentBlock)); // Сохраняем текущий блок
                    currentBlock.clear(); // Начинаем новый блок
                }
            }
        }

        // Добавляем последний блок, если он не пустой
        if (!currentBlock.isEmpty()) {
            dataBlocks.add(new ArrayList<>(currentBlock));
        }

        workbook.close();
        file.close();
        return dataBlocks;
    }

    private DefaultItem processDataBlock(Row row) {
        DefaultItem item = new DefaultItem();

        // Читаем A-F (индексы 0-5)
        item.setItemNo(getIntegerValue(row.getCell(0)));                // A
        item.setOriginalName(getCellValue(row.getCell(1)));             // B
        item.setAlterNameRus(getCellValue(row.getCell(2)));             // C
        item.setSize(getCellValue(row.getCell(3)));                     // D
        item.setMarking(getCellValue(row.getCell(4)));                  // E
        item.setQuantityInBox(getCellValue(row.getCell(5)));            // F
        item.setOrder(getCellValue(row.getCell(6)));                    // G
        item.setAlterImagePath(getCellValue(row.getCell(7)));           // H

        // Логируем значения для каждой ячейки
        logger.debug("Читаем строку -> ItemNo: {}, OriginalName: {}, AlterNameRus: {}, Size: {}, Marking: {}, QuantityInBox: {}, Order: {}, AlterImagePath: {}",
                item.getItemNo(), item.getOriginalName(), item.getAlterNameRus(), item.getSize(),
                item.getMarking(), item.getQuantityInBox(), item.getOrder(), item.getAlterImagePath());

        if (isEmptyItem(item)) {
            logger.info("Элемент пустой, пропускаем");
            return null;
        }

        return item;
    }


    private boolean isEmptyItem(DefaultItem item) {
        // Проверяем, что все поля пустые
        return (item.getItemNo() == null) &&
                (item.getOriginalName() == null || item.getOriginalName().trim().isEmpty()) &&
                (item.getAlterNameRus() == null || item.getAlterNameRus().trim().isEmpty()) &&
                (item.getSize() == null || item.getSize().trim().isEmpty()) &&
                (item.getQuantityInBox() == null || item.getQuantityInBox().trim().isEmpty()) &&
                (item.getMarking() == null || item.getMarking().trim().isEmpty()) &&
                (item.getAlterImagePath() == null || item.getAlterImagePath().trim().isEmpty()) &&
                (item.getOrder() == null || item.getOrder().trim().isEmpty());
    }

    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf((int) cell.getNumericCellValue()); // Преобразуем в int, если число
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

    private Integer getIntegerValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            return (int) cell.getNumericCellValue(); // Преобразуем число в Integer
        }
        if (cell.getCellType() == CellType.STRING) {
            try {
                return Integer.parseInt(cell.getStringCellValue().trim());
            } catch (NumberFormatException e) {
                return null;
            }
        }
        return null;
    }
}
