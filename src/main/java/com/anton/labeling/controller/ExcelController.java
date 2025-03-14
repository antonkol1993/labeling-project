package com.anton.labeling.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
public class ExcelController {

    @PostMapping("/upload")
    public String uploadExcel(@RequestParam("file") MultipartFile file) {
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println();
            }
            return "Файл успешно обработан!";
        } catch (IOException e) {
            return "Ошибка при обработке файла: " + e.getMessage();
        }
    }
}
