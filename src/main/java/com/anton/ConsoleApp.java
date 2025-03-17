package com.anton;

import java.io.FileOutputStream;
import java.io.IOException;

import com.anton.labeling.objects.ItemLargeBox;
import com.anton.labeling.service.largeBox.CardCreator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ConsoleApp {
    public static void main(String[] args) {
        ItemLargeBox item = new ItemLargeBox();
        item.setName("Саморезы гипс/металл");
        item.setSize("3.5x25");
        item.setMarking("YZP");
        item.setQuantityInBox("1000");
        item.setOrder("2155695PL");
        item.setNameAndSize(item.getName() + "\t" + item.getSize());

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Card");

        CardCreator cardCreator = new CardCreator();
//        for (int i = 1; i < 90; i += 4) {
            try {
                cardCreator.createCard(workbook, sheet, item, 1, 1);
            } catch (IOException e) {
                e.printStackTrace();
            }
//        }

        // Сохранение файла и закрытие workbook
        try (FileOutputStream fileOut = new FileOutputStream("ExcelCard.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        System.out.println("Excel файл успешно создан: ExcelCard.xlsx");
    }
}
