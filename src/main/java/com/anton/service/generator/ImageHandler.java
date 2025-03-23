package com.anton.service.generator;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;

public class ImageHandler {

    public static void addImageToSheet(Workbook workbook, Sheet sheet, String imagePath,
                                       int startRow, int startCol, int endRow, int endCol) throws IOException {
        // Загружаем изображение в массив байтов
        byte[] imageBytes;
        try (InputStream inputStream = new FileInputStream(imagePath)) {
            imageBytes = IOUtils.toByteArray(inputStream);
        }

        // Определяем тип изображения
        int pictureType;
        if (imagePath.toLowerCase().endsWith(".png")) {
            pictureType = Workbook.PICTURE_TYPE_PNG;
        } else if (imagePath.toLowerCase().endsWith(".jpeg") || imagePath.toLowerCase().endsWith(".jpg")) {
            pictureType = Workbook.PICTURE_TYPE_JPEG;
        } else {
            throw new IllegalArgumentException("Поддерживаются только PNG и JPEG изображения.");
        }

        int pictureIdx = workbook.addPicture(imageBytes, pictureType);

        // Загружаем изображение для получения его размеров
        BufferedImage bufferedImage = ImageIO.read(new ByteArrayInputStream(imageBytes));
        int originalWidth = bufferedImage.getWidth();
        int originalHeight = bufferedImage.getHeight();

        // Вычисляем размеры ячейки в пикселях
        float cellWidthPx = 0;
        for (int col = startCol; col <= endCol; col++) {
            cellWidthPx += sheet.getColumnWidthInPixels(col);
        }

        float cellHeightPx = 0;
        for (int row = startRow; row <= endRow; row++) {
            Row sheetRow = sheet.getRow(row);
            if (sheetRow != null) {
                cellHeightPx += sheetRow.getHeightInPoints() * 1.33f;
            }
        }

        // Масштабируем изображение так, чтобы оно полностью влезало
        double scaleX = cellWidthPx / originalWidth;
        double scaleY = cellHeightPx / originalHeight;

        int imageWidth = (int) (originalWidth);
        int imageHeight = (int) (originalHeight);

        // Центрирование изображения в ячейке после масштабирования
        double offsetX = (cellWidthPx - imageWidth) / 2;
        double offsetY = (cellHeightPx - imageHeight) / 2;

        // Конвертация в EMU (Excel Measurement Units)
        int dx1 = offsetX < 0 ? (int) (-offsetX * 9525) : (int) (offsetX * 9525);
        int dy1 = offsetY < 0 ? (int) (-offsetY * 9525) : (int) (offsetY * 9525);

        // Создаём рисунок
        if (workbook instanceof XSSFWorkbook) {
            XSSFDrawing drawing = ((XSSFWorkbook) workbook).getSheetAt(0).createDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, 0, 0, startCol, startRow, endCol + 1, endRow + 1);
            Picture picture = drawing.createPicture(anchor, pictureIdx);

            // Устанавливаем масштабированное изображение
            picture.resize();
        }
    }
}
