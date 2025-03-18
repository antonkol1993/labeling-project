package com.anton;

import com.anton.labeling.objects.ItemLargeBox;
import com.anton.service.DynamicExcelGenerator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ConsoleApp {
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
            XSSFSheet sheet = workbook.createSheet("Sheet1");

//            for (int j = 0; j < 5; j++) {
//                int temp = startCol;

                for (int i = 0; i < 30; i++) {
                    DynamicExcelGenerator generator = new DynamicExcelGenerator(sheet, startRow, startCol, workbook);
                    generator.addCard(item);
                    startCol += 4; // Оставляем 1 столбец между карточками
                }
//                startRow += 12;
//                startCol=temp;
//            }
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            }

            System.out.println("Файл создан: " + fileName);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
