package com.anton;

import java.io.IOException;
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
