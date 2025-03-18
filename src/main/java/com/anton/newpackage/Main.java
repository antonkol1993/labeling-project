package com.anton.newpackage;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        ItemLargeBox item = new ItemLargeBox();
        item.setName("Саморезы гипс/металл");
        item.setSize("3.5x25");
        item.setMarking("YZP");
        item.setQuantityInBox("1000");
        item.setOrder("2155695PL");
        item.setNameAndSize(item.getName() + "\t" + item.getSize());


        String fileName = "output.xlsx";
        int startRow = 2;
        int startCol = 2;

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            for (int i = 0; i < 30; i++) {
                DynamicExcelGenerator generator = new DynamicExcelGenerator(sheet, startRow, startCol);
                generator.addCard(item);
                startCol += 4; // Оставляем 1 столбец между карточками
            }

            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            }

            System.out.println("Файл создан: " + fileName);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
