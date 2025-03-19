package com.anton.service.reader;

import com.anton.labeling.objects.ItemLargeBox;

import java.io.IOException;
import java.util.List;

public class ReaderMain {
    public static void main(String[] args) {
        try {
            ExcelDataReader reader = new ExcelDataReader();
            List<List<ItemLargeBox>> blocks = reader.readExcel("excel-example/DataFromInvoice .xlsx");

            System.out.println("Количество блоков: " + blocks.size());
            for (int i = 0; i < blocks.size(); i++) {
                System.out.println("Блок " + (i + 1) + " содержит " + blocks.get(i).size() + " элементов");

            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}