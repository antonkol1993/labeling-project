package com.anton;

import com.anton.labeling.service.largeBox.CardCreator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class ConsoleApp {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Card");

// Создаем экземпляр CardCreator и вызываем метод createCard
        CardCreator cardCreator = new CardCreator();
        cardCreator.createCard(workbook, sheet);
    }
}
