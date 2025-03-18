package com.anton;

import com.anton.labeling.objects.ItemLargeBox;
import com.anton.labeling.service.irrelevantOld.CardCreator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class ConsoleApp {
    public static void main(String[] args) throws IOException {
        ItemLargeBox item = new ItemLargeBox();
        item.setName("Саморезы гипс/металл");
        item.setSize("3.5x25");
        item.setMarking("YZP");
        item.setQuantityInBox("1000");
        item.setOrder("2155695PL");
        item.setNameAndSize(item.getName() + "\t" + item.getSize());

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Card");

// Создаем экземпляр CardCreator и вызываем метод createCard
        CardCreator cardCreator = new CardCreator();
        cardCreator.createCard(workbook, sheet, item);
    }
}
