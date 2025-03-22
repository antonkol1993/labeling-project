package com.anton.service.mapper;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelMapper {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "excel-example/China14 invoices/25HS10047P-PI  Final 3.13.xlsx";
        String propertiesFilePath = "src/main/resources/mapping_item-invoice.properties";

        Map<String, String> mapping = loadProperties(propertiesFilePath);
        Map<String, String> result = processExcelFile(excelFilePath, mapping);

        result.forEach((key, value) -> System.out.println(key + " = " + value));
    }

    private static Map<String, String> loadProperties(String filePath) throws IOException {
        Properties properties = new Properties();
        try (InputStream input = new FileInputStream(filePath)) {
            properties.load(input);
        }
        Map<String, String> map = new HashMap<>();
        for (String key : properties.stringPropertyNames()) {
            map.put(properties.getProperty(key), key);
        }
        return map;
    }

    private static Map<String, String> processExcelFile(String filePath, Map<String, String> mapping) throws IOException {
        Map<String, String> result = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (isMergedInRange(sheet, cell, 2, 14)) { // Проверяем, объединена ли ячейка в колонках C:O (индексы 2–14)
                        int rowIndex = row.getRowNum();
                        Cell bCell = row.getCell(1); // Колонка B (индекс 1)
                        Cell above = (rowIndex > 0) ? sheet.getRow(rowIndex - 1).getCell(1) : null;
                        Cell below = (rowIndex < sheet.getLastRowNum()) ? sheet.getRow(rowIndex + 1).getCell(1) : null;

                        if (isValid(above, below, bCell)) {
                            String value = getMergedCellValue(sheet, cell).trim();
                            if (mapping.containsKey(value)) {
                                result.put(mapping.get(value), value);
                            }
                        }
                    }
                }
            }
        }

        return result;
    }

    private static boolean isValid(Cell above, Cell below, Cell bCell) {
        boolean bCellEmpty = isCellEmpty(bCell);
        boolean aboveIsNumber = isNumeric(above);
        boolean belowIsNumber = isNumeric(below);

        if (aboveIsNumber && belowIsNumber) {
            return bCellEmpty && (getNumericValue(above) + 1 == getNumericValue(below));
        } else if (!aboveIsNumber && belowIsNumber) {
            return bCellEmpty && getNumericValue(below) == 1;
        }
        return false;
    }

    private static boolean isNumeric(Cell cell) {
        return cell != null && cell.getCellType() == CellType.NUMERIC;
    }

    private static double getNumericValue(Cell cell) {
        return cell.getNumericCellValue();
    }

    private static boolean isCellEmpty(Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK;
    }

    private static String getMergedCellValue(Sheet sheet, Cell cell) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                Row firstRow = sheet.getRow(range.getFirstRow());
                Cell firstCell = firstRow.getCell(range.getFirstColumn());
                return getCellValue(firstCell);
            }
        }
        return getCellValue(cell);
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            default -> "";
        };
    }

    private static boolean isMergedInRange(Sheet sheet, Cell cell, int colStart, int colEnd) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex()) &&
                    cell.getColumnIndex() >= colStart && cell.getColumnIndex() <= colEnd) {
                return true;
            }
        }
        return false;
    }

}
