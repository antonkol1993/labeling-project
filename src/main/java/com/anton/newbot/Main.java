package com.anton.newbot;

import com.anton.labeling.objects.ItemLargeBox;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Card");

            ItemLargeBox item = new ItemLargeBox();
            item.setName("Саморезы гипс/металл");
            item.setSize("3.5x25");
            item.setMarking("YZP");
            item.setQuantityInBox("1000");
            item.setOrder("2155695PL");
            item.setNameAndSize(item.getName() + "\t" + item.getSize());


            CardCreator cardCreator = new CardCreator(workbook, sheet);
            cardCreator.createCard(item, 2, 1);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
