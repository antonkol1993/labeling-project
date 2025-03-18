package com.anton.newpackage;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        String fileName = "output.xlsx";
        int startRow = 2;
        int startCol = 2;

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            for (int i = 0; i < 30; i++) {
                DynamicExcelGenerator generator = new DynamicExcelGenerator(sheet, startRow, startCol);
                generator.addCard();
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
