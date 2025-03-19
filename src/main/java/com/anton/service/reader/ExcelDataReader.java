package com.anton.service.reader;


import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ExcelDataReader {

    private final List<List<ItemLargeBox>> dataBlocks = new ArrayList<>();
    private List<ItemLargeBox> currentBlock = new ArrayList<>();

    public List<List<ItemLargeBox>> readExcel(String filePath) throws IOException {
        FileInputStream file = new FileInputStream(new File(filePath));
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);

        boolean headersProcessed = false;

        for (Row row : sheet) {
            if (row.getRowNum() < 2) continue; // Пропускаем заголовки (читаем с 3 строки)

            if (!headersProcessed) {
                processHeaders(row);
                headersProcessed = true;
                continue;
            }

            // Обрабатываем строку
            ItemLargeBox item = processDataBlock(row);

            if (item != null) {
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

    public List<List<ItemLargeBox>> getDataBlocks() {
        return dataBlocks;
    }

    private void processHeaders(Row row) {
        System.out.println("Заголовки: ");
        for (Cell cell : row) {
            System.out.print(cell.getCellType() + " | ");
        }
        System.out.println("\n-------------------------");
    }

    private ItemLargeBox processDataBlock(Row row) {
        ItemLargeBox item = new ItemLargeBox();

        // Читаем колонку A (индекс 0) - invoiceItemNumber
        item.setInvoiceItemNumber(getIntegerValue(row.getCell(0)));

        // Читаем B-F (индексы 1-5)
        item.setName(getCellValue(row.getCell(1)));
        item.setSize(getCellValue(row.getCell(2)));
        item.setQuantityInBox(getCellValue(row.getCell(3)));
        item.setMarking(getCellValue(row.getCell(4)));
        item.setOrder(getCellValue(row.getCell(5)));

        item.setNameAndSize(item.getName() + "\n" + item.getSize());

        if (isEmptyItem(item)) {
            return null;
        }

        return item;
    }

    private boolean isEmptyItem(ItemLargeBox item) {
        return (item.getInvoiceItemNumber() == null) &&
                (item.getName() == null || item.getName().trim().isEmpty()) &&
                (item.getSize() == null || item.getSize().trim().isEmpty()) &&
                (item.getQuantityInBox() == null || item.getQuantityInBox().trim().isEmpty()) &&
                (item.getMarking() == null || item.getMarking().trim().isEmpty()) &&
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
