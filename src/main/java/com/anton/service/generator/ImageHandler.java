package com.anton.service.generator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.io.IOUtils;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ImageHandler {
    public static void addImageToSheet(Workbook workbook, Sheet sheet, String imagePath,
                                       int row1, int col1, int row2, int col2,
                                       int dx1, int dy1) throws IOException {
        InputStream inputStream = new FileInputStream(imagePath);
        byte[] imageBytes = IOUtils.toByteArray(inputStream);
        inputStream.close();

        int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);

        if (workbook instanceof XSSFWorkbook) {
            XSSFDrawing drawing = ((XSSFWorkbook) workbook).getSheetAt(0).createDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, 0, 0, col1, row1, col2, row2);
            Picture picture = drawing.createPicture(anchor, pictureIdx);
            picture.resize();
        } else if (workbook instanceof HSSFWorkbook) {
            HSSFPatriarch drawing = ((HSSFWorkbook) workbook).getSheetAt(0).createDrawingPatriarch();
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 255, (short) col1, row1, (short) col2, row2);
            drawing.createPicture(anchor, pictureIdx);
        }
    }
}
