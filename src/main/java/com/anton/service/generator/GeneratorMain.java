package com.anton.service.generator;

import com.anton.labeling.objects.ItemLargeBox;
import com.anton.service.reader.ExcelDataReader;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class GeneratorMain {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Карточки");
        DynamicExcelGenerator generator = new DynamicExcelGenerator(workbook, sheet);

        ExcelDataReader reader = new ExcelDataReader();
        List<List<ItemLargeBox>> dataBlocks = reader.readExcel("excel-example/DataFromInvoice .xlsx");

        generator.generateCardsFromBlocks(dataBlocks);

        try (FileOutputStream fileOut = new FileOutputStream("output.xlsx")) {
            workbook.write(fileOut);
        }
        workbook.close();
    }
}
