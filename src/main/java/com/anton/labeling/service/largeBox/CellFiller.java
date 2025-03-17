package com.anton.labeling.service.largeBox;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public class CellFiller {
    private final XSSFWorkbook workbook;
    private final XSSFSheet sheet;

    public CellFiller(XSSFWorkbook workbook, XSSFSheet sheet) {
        this.workbook = workbook;
        this.sheet = sheet;
    }

    public void fillCellsWithData(ItemLargeBox item, int startRow, int startCol) {
        // Определяем стили
        CellStyle centerBold11 = CardStyle.createCenteredBoldStyle(workbook, (short) 11);
        CellStyle centerBold10 = CardStyle.createCenteredBoldStyle(workbook, (short) 10);
        CellStyle leftBold10 = CardStyle.createLeftBoldStyle(workbook, (short) 10);

        Object[][] values = {
                {3, 1, 3, item.getName() + "\n" + item.getSize(), centerBold11},  // 3(1-3)
                {4, 1, 1, "Marking", leftBold10},
                {5, 1, 1, "РАЗМЕР/Size", leftBold10},
                {7, 1, 1, "Кол-во в упак/шт.", leftBold10},
                {8, 1, 1, "Вес упак Кг/Kgs", leftBold10},
                {10, 1, 1, "ORDER:", leftBold10},
                {4, 2, 3, item.getMarking(), centerBold10},  // 4(2-3)
                {5, 2, 3, item.getSize(), centerBold10},     // 5(2-3)
                {7, 2, 2, "Шт / PCS", centerBold10},        // 7(2)
                {7, 3, 3, "", leftBold10},                  // 7(3)
                {8, 3, 3, "Кг/Kgs", centerBold10},          // 8(3)
                {8, 3, 3, "", leftBold10},                  // 8(3)
                {9, 2, 3, "Сделано в КНР", leftBold10},     // 9(2-3)
                {10, 2, 3, item.getOrder(), leftBold10}     // 10(2-3)
        };

        for (Object[] entry : values) {
            int rowIdx = startRow + (int) entry[0] - 1;  // Смещение строки
            int colIdx = startCol + (int) entry[1] - 1;  // Смещение столбца
            int colSpan = (int) entry[2]; // Кол-во объединяемых колонок
            String value = (String) entry[3];
            CellStyle style = (CellStyle) entry[4];

            Row row = sheet.getRow(rowIdx);
            if (row == null) {
                row = sheet.createRow(rowIdx);
            }

            Cell cell = row.createCell(colIdx);
            cell.setCellValue(value);
            cell.setCellStyle(style);

            // Объединение ячеек, если colSpan > 1
            if (colSpan > 1) {
                CellRangeAddress newRegion = new CellRangeAddress(rowIdx, rowIdx, colIdx, colIdx + colSpan - 1);
                if (!isOverlapping(sheet, newRegion)) {
                    sheet.addMergedRegion(newRegion);
                }
            }
        }
    }

    // Метод проверки, пересекается ли объединение с уже существующими
    private boolean isOverlapping(XSSFSheet sheet, CellRangeAddress newRegion) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress existing : mergedRegions) {
            if (existing.intersects(newRegion)) {
                return true;
            }
        }
        return false;
    }

}