package com.anton;

import com.anton.objects.DefaultItem;
import com.anton.service.generator.DynamicExcelGeneratorLargeBoxes;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Scanner;

public class ConsoleApp {
    public static void main(String[] args) throws IOException {
        Scanner scanner = new Scanner(System.in);
        System.out.println("Выберите формат файла: 1 - .xlsx, 2 - .xls");
        int choice = scanner.nextInt();
        scanner.nextLine(); // Очистка буфера после nextInt

        String fileName = choice == 1 ? "output.xlsx" : "output.xls";
        boolean isXSSF = choice == 1;

//        DataReaderDefault dataReaderDefault = new DataReaderDefault();
//        List<List<DefaultItem>> dataBlocks = dataReaderDefault.readExcel("excel-example/DataFromInvoice .xlsx");
//
//
//
//        try (Workbook workbook = isXSSF ? new XSSFWorkbook() : new HSSFWorkbook()) {
//            Sheet sheet = workbook.createSheet("Sheet1");
//
//            DynamicExcelGeneratorLargeBoxes generator = new DynamicExcelGeneratorLargeBoxes(workbook, sheet);
//            generator.generateCardsFromBlocks(dataBlocks);
//
//            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
//                workbook.write(fileOut);
//            }
//
//            System.out.println("Файл создан: " + fileName);
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
    }
}
