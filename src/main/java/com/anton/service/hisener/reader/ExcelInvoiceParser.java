package com.anton.service.hisener.reader;

import com.anton.objects.InvoiceItemData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.text.DecimalFormat;
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
                            Objects.toString(item.getName(), "null"),
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
        DecimalFormat df = new DecimalFormat("#0.00"); // Формат с 2 знаками после запятой
        double totalSum = parsedData.stream()
                .flatMap(List::stream) // Разворачиваем группы в единый поток элементов
                .mapToDouble(item -> Objects.requireNonNullElse(item.getTotal(), 0.0)) // Преобразуем в double, заменяя null на 0
                .sum();

        System.out.println("Общая сумма Total: " + df.format(totalSum));
    }


    public static List<List<InvoiceItemData>> parseExcel(String filePath) throws IOException {
        String groupName = "";
        List<List<InvoiceItemData>> groupedData = new ArrayList<>();
        List<InvoiceItemData> currentGroup = null;
        String currentProformaNo = null;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {  // Автоматически определяет формат (XLS/XLSX)


            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator(); // Создаем FormulaEvaluator для вычисления формул
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {

// Проверяем строку 9 на наличие объединенных ячеек для proformaNo (столбцы N и O)
                boolean isXLS = filePath.toLowerCase().endsWith(".xls");

                int proformaRowIndex = isXLS ? 7 : 8; // В .xls ищем на строке 8 (индекс 7), в .xlsx на строке 9 (индекс 8)

                if (row.getRowNum() == proformaRowIndex) {
                    Cell proformaCell = row.getCell(13); // Столбец N (индексация с 0)
                    if (isMergedInRange(sheet, proformaCell, 13, 14)) {
                        currentProformaNo = getCellValue(proformaCell);
                        System.out.println("✅ Найден proformaNo: " + currentProformaNo);
                    }
                }

                // Проверяем наличие данных в колонке B
                Cell elementCell = row.getCell(1);
                if (isNumeric(elementCell)) {
                    // Проверка строки выше на наличие объединённых ячеек с C по O
                    Row prevRow = sheet.getRow(row.getRowNum() - 1); // Строка выше
                    if (prevRow != null) {
                        // Проверяем, есть ли объединение в строке выше (C-O)
                        Cell prevElementCell = prevRow.getCell(2); // Проверка в строке выше с C по O
                        if (isMergedInRange(sheet, prevElementCell, 2, 14)) {
                            currentGroup = new ArrayList<>(); // Создаём новую группу
                            groupedData.add(currentGroup);     // Добавляем её в общий список
                            System.out.println("Создан новый список для группы.");
                            // Получаем имя из объединенной ячейки C-O
                            groupName = getMergedCellValue(sheet, prevElementCell).trim();

                        }
                    }

                    // Создаем объект InvoiceItemData
                    InvoiceItemData item = new InvoiceItemData();
                    item.setName(groupName);
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

                    // Добавляем объект в текущую группу
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
