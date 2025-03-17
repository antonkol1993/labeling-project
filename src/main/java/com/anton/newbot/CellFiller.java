package com.anton.newbot;


import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.*;
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
        CellStyle centerBold11 = CardStyle.createCenteredBoldStyle(workbook, (short) 11);
        CellStyle centerBold10 = CardStyle.createCenteredBoldStyle(workbook, (short) 10);
        CellStyle leftBold10 = CardStyle.createLeftBoldStyle(workbook, (short) 10);

        Object[][] values = {
                {1, 1, 3, "", null},
                {2, 1, 3, "", null},
                {3, 1, 3, item.getName() + "\n" + item.getSize(), centerBold11},
                {4, 1, 1, "Marking", leftBold10},
                {4, 2, 3, item.getMarking(), centerBold10},
                {5, 1, 1, "РАЗМЕР/Size", leftBold10},
                {5, 2, 3, item.getSize(), centerBold10},
                {7, 1, 1, "Кол-во в упак/шт.", leftBold10},
                {7, 2, 2, item.getQuantityInBox(), centerBold10},
                {7, 3, 3, "Шт / PCS", leftBold10},  // Сдвиг на одну ячейку вправо (было 2 -> стало 3)
                {8, 1, 1, "Вес упак Кг/Kgs", leftBold10},
                {8, 3, 3, "Кг/Kgs", leftBold10},  // Сдвиг на одну ячейку вправо (было 2 -> стало 3)
                {9, 2, 3, "Сделано в КНР", leftBold10},
                {10, 1, 1, "ORDER:", leftBold10},
                {10, 2, 3, item.getOrder(), leftBold10}
        };

        for (Object[] value : values) {
            int rowIdx = startRow + (int) value[0] - 1;
            int colIdx = startCol + (int) value[1] - 1;
            int mergeCols = (int) value[2];
            String text = (String) value[3];
            CellStyle style = (CellStyle) value[4];

            // Создание строки, если её нет
            Row row = sheet.getRow(rowIdx);
            if (row == null) row = sheet.createRow(rowIdx);

            // Создание ячеек
            for (int i = 0; i < mergeCols; i++) {
                Cell cell = row.createCell(colIdx + i);
                if (i == 0) {  // Устанавливаем текст только в первую ячейку объединения
                    cell.setCellValue(text);
                    if (style != null) cell.setCellStyle(style);
                }
            }
        }
    }
}

