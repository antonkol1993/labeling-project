package com.anton.labeling.service.largeBox;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellFiller {
    private final XSSFWorkbook workbook;
    private final XSSFSheet sheet;

    public CellFiller(XSSFWorkbook workbook, XSSFSheet sheet) {
        this.workbook = workbook;
        this.sheet = sheet;
    }

    public void fillCellsWithData(ItemLargeBox item, int startRow, int startCol) {
        Object[][] values = {
                {3, 1, 3, item.getName() + "\n" + item.getSize()}, // 3(1-3)
                {4, 1, 1, "Marking"},
                {5, 1, 1, "РАЗМЕР/Size"},
                {6, 1, 1, "Кол-во/Q-ty"},
                {7, 1, 1, "Кол-во в упак/шт."},
                {8, 1, 1, "Вес упак Кг/Kgs"},
                {9, 1, 1, ""},  // пустая строка
                {10, 1, 1, "ORDER:"},
                {4, 2, 3, item.getMarking()},  // 4(2-3)
                {5, 2, 3, item.getSize()},     // 5(2-3)
                {7, 3, 3, "Шт / PCS"},        // 7(3)
                {8, 3, 3, "Кг/Kgs"},          // 8(3)
                {9, 2, 3, "Сделано в КНР"},   // 9(2-3)
                {10, 2, 3, item.getOrder()}   // 10(2-3)
        };

        for (Object[] val : values) {
            int rowIdx = startRow + (int) val[0] - 1;
            int colStart = startCol + (int) val[1] - 1;
            int colEnd = startCol + (int) val[2] - 1;
            String text = val[3].toString();

            Row row = sheet.getRow(rowIdx);
            if (row == null) row = sheet.createRow(rowIdx);

            for (int col = colStart; col <= colEnd; col++) {
                Cell cell = row.getCell(col);
                if (cell == null) cell = row.createCell(col);
                cell.setCellValue(text);
                cell.setCellStyle(CardStyle.createBorderedCellStyle(workbook, startRow, startCol, rowIdx, col));
            }

            // Проверяем, нужно ли объединение ячеек
            CellRangeAddress newRegion = new CellRangeAddress(rowIdx, rowIdx, colStart, colEnd);
            if (newRegion.getNumberOfCells() > 1) {
                boolean alreadyMerged = false;
                for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                    CellRangeAddress existingRegion = sheet.getMergedRegion(i);
                    if (existingRegion.getFirstRow() == newRegion.getFirstRow() &&
                            existingRegion.getLastRow() == newRegion.getLastRow() &&
                            existingRegion.getFirstColumn() == newRegion.getFirstColumn() &&
                            existingRegion.getLastColumn() == newRegion.getLastColumn()) {
                        alreadyMerged = true;
                        break;
                    }
                }
                if (!alreadyMerged) {
                    sheet.addMergedRegion(newRegion);
                }
            }
        }
    }



}