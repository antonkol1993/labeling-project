package com.anton.labeling.controller;

import com.anton.labeling.objects.ItemLargeBox;

import java.io.IOException;
import java.util.List;

public class Main {
    public static void main(String[] args) {

        try {
            ExcelDataReader excelDataReader = new ExcelDataReader();
            List<ItemLargeBox> items = excelDataReader.readExcel("excel-example/DataFromInvoice .xlsx");

            for (ItemLargeBox item : items) {
                System.out.println(item.getNameAndSize() + " | " + item.getQuantityInBox());
            }
            System.out.println(excelDataReader.totalEmptyRows());

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}