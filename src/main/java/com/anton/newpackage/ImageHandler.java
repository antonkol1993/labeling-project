package com.anton.newpackage;

import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ImageHandler {
    public static void addImageToSheet(XSSFWorkbook workbook, XSSFSheet sheet, String imagePath,
                                       int row1, int col1, int row2, int col2,
                                       int dx1,int dy1) throws IOException {
        InputStream inputStream = new FileInputStream(imagePath);
        byte[] imageBytes = IOUtils.toByteArray(inputStream);
        int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
        inputStream.close();

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = new XSSFClientAnchor();

        anchor.setCol1(col1);
        anchor.setRow1(row1);
        anchor.setCol2(col2);
        anchor.setRow2(row2);
        anchor.setDx1(dx1);
        anchor.setDy1(dy1);

        Picture picture = drawing.createPicture(anchor, pictureIdx);
        picture.resize();
    }
}

