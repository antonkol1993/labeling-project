package com.anton.service.mapper;

import com.anton.labeling.objects.InvoiceItemData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelInvoiceParser {

    public static void main(String[] args) throws IOException {
        String filePath = "excel-example/China14 invoices/25HS10047P-PI  Final 3.13.xlsx";
        List<List<InvoiceItemData>> parsedData = ExcelInvoiceParser.parseExcel(filePath);

        parsedData.forEach(group -> {
            System.out.println("=== Новая группа ===");
            group.forEach(item -> System.out.println(
                    String.join(" | ",
                            Objects.toString(item.getProformaNo(), "null"),
                            Objects.toString(item.getElementNumber(), "null"),
                            Objects.toString(item.getSize(), "null"),
                            Objects.toString(item.getPartNo(), "null"),
                            Objects.toString(item.getFinish(), "null"),
                            Objects.toString(item.getBox(), "null"),
                            Objects.toString(item.getCtn(), "null"),
                            Objects.toString(item.getCtnPltCtn(), "null"),
                            Objects.toString(item.getKgsUn(), "null"),
                            Objects.toString(item.getTotalKgs(), "null"),
                            Objects.toString(item.getQuantity(), "null"),
                            Objects.toString(item.getUnit(), "null"),
                            Objects.toString(item.getUnitPrice(), "null"),
                            Objects.toString(item.getTotal(), "null")
                    )
            ));
        });
    }


    public static List<List<InvoiceItemData>> parseExcel(String filePath) throws IOException {
        List<List<InvoiceItemData>> groupedData = new ArrayList<>();
        List<InvoiceItemData> currentGroup = null;
        String currentProformaNo = null;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator(); // Создаем FormulaEvaluator для вычисления формул
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {

// Проверяем строку 9 на наличие объединенных ячеек для proformaNo (столбцы N и O)
                if (row.getRowNum() == 8) { // Строка 9 (индексация с 0)
                    Cell proformaCell = row.getCell(13); // Столбец N (индексация с 0)
                    if (isMergedInRange(sheet, proformaCell, 13, 14)) {
                        currentProformaNo = getCellValue(proformaCell);
                        System.out.println("✅ Найден proformaNo: " + currentProformaNo);
                    }
                }

                // Парсим строку, если в колонке B есть число
                Cell elementCell = row.getCell(1);
                if (isNumeric(elementCell)) {
                    if (currentGroup == null) {
                        currentGroup = new ArrayList<>();
                        groupedData.add(currentGroup);
                    }

                    InvoiceItemData item = new InvoiceItemData();
                    item.setProformaNo(currentProformaNo);

                    item.setElementNumber(getIntegerValueOrNull(elementCell));
                    item.setSize(getCellValue(row.getCell(2)));
                    item.setPartNo(getCellValue(row.getCell(3)));
                    item.setFinish(getCellValue(row.getCell(5)));
                    item.setBox(getIntegerValueOrNull(row.getCell(6)));
                    item.setCtn(getCellValue(row.getCell(7)));
                    item.setCtnPltCtn(getCellValue(row.getCell(8)));
                    item.setKgsUn(getDoubleValueOrNull(row.getCell(9), formulaEvaluator));
                    item.setTotalKgs(getDoubleValueOrNull(row.getCell(10), formulaEvaluator));
                    item.setQuantity(getDoubleValueOrNull(row.getCell(11), formulaEvaluator));
                    item.setUnit(getCellValue(row.getCell(12)));
                    item.setUnitPrice(getDoubleValueOrNull(row.getCell(13), formulaEvaluator));

                    // Проверяем total (O-колонка)
                    Cell totalCell = row.getCell(14);
                    Double totalValue = getDoubleValueOrNull(totalCell, formulaEvaluator);
                    if (totalValue == null) {
                        System.out.println("⚠️ Проблема с total в строке " + row.getRowNum() + ": " + getCellValue(totalCell));
                    }
                    item.setTotal(totalValue);

                    currentGroup.add(item);
                }
            }
        }
        return groupedData;
    }

    // Проверяет, является ли ячейка частью объединенной ячейки
    private static boolean isMergedInRange(Sheet sheet, Cell cell, int colStart, int colEnd) {
        if (cell == null) return false;
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex()) &&
                    cell.getColumnIndex() >= colStart && cell.getColumnIndex() <= colEnd) {
                return true;
            }
        }
        return false;
    }

    // Проверяет, является ли ячейка числом
    private static boolean isNumeric(Cell cell) {
        if (cell == null) return false;
        String value = getCellValue(cell);
        if (value == null || value.isEmpty()) return false;

        try {
            Double.parseDouble(value);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    // Возвращает значение ячейки в виде строки
    private static String getCellValue(Cell cell) {
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            default -> null;
        };
    }

    // Возвращает Double значение ячейки или null, если ячейка пуста или не число (с учетом формулы)
    private static Double getDoubleValueOrNull(Cell cell, FormulaEvaluator formulaEvaluator) {
        if (cell == null) return null;

        // Если ячейка имеет формулу, вычисляем ее
        if (cell.getCellType() == CellType.FORMULA) {
            return formulaEvaluator.evaluate(cell).getNumberValue();
        }

        // Если ячейка не имеет формулы, возвращаем числовое значение
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }

        return null;
    }

    // Возвращает Integer значение ячейки или null, если ячейка пуста или не число
    private static Integer getIntegerValueOrNull(Cell cell) {
        Double value = getDoubleValueOrNull(cell, null);
        return (value != null) ? value.intValue() : null;
    }
}
