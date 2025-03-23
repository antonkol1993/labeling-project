package com.anton.service.generator;

import com.anton.labeling.objects.InvoiceItemData;
import com.anton.labeling.objects.LabelLargeBox;
import com.anton.service.mapper.LabelLargeBoxMapper;
import com.anton.service.reader.ExcelDataReader;
import com.anton.service.reader.ExcelInvoiceParser;
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
        DynamicExcelGeneratorLargeBoxes generator = new DynamicExcelGeneratorLargeBoxes(workbook, sheet);

        ExcelDataReader reader = new ExcelDataReader();
//        List<List<LabelLargeBox>> dataBlocks = reader.readExcel("excel-example/DataFromInvoice .xlsx");
//
//        generator.generateCardsFromBlocks(dataBlocks);
        String filePath = "excel-example/China14 invoices/25HS10047P-PI  Final 3.13.xlsx";
        List<List<InvoiceItemData>> parsedData = ExcelInvoiceParser.parseExcel(filePath);

        LabelLargeBoxMapper mapper = new LabelLargeBoxMapper();
        List<List<LabelLargeBox>> mappedData = mapper.mapData(parsedData);
        generator.generateCardsFromBlocks(mappedData);

        try (FileOutputStream fileOut = new FileOutputStream("output.xlsx")) {
            workbook.write(fileOut);
        }
        workbook.close();
    }
}
