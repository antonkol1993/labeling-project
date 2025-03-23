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

        // Добавляем изображение в Workbook
        int pictureIdx = workbook.addPicture(imageBytes, pictureType);

        // Загружаем изображение для получения его размеров
        BufferedImage bufferedImage = ImageIO.read(new ByteArrayInputStream(imageBytes));
        int originalWidth = bufferedImage.getWidth();
        int originalHeight = bufferedImage.getHeight();

        // Получаем размеры ячейки в пикселях
        float cellWidthPx = 0;
        for (int col = startCol; col <= endCol; col++) {
            cellWidthPx += sheet.getColumnWidthInPixels(col);
        }

        float cellHeightPx = 0;
        for (int row = startRow; row <= endRow; row++) {
            Row sheetRow = sheet.getRow(row);
            if (sheetRow != null) {
                cellHeightPx += sheetRow.getHeightInPoints() * 1.33f;  // Преобразуем высоту в пиксели
            }
        }

        // Рассчитываем коэффициент масштабирования, чтобы изображение не превышало 85% от размера ячейки
        double scaleX = cellWidthPx * 0.85 / originalWidth;  // 85% от ширины ячейки
        double scaleY = cellHeightPx * 0.85 / originalHeight;  // 85% от высоты ячейки

        // Масштабируем по большей стороне изображения
        double scale = Math.min(scaleX, scaleY);

        // Если изображение меньше ячейки, оставляем scale = 1
        if (originalWidth < cellWidthPx && originalHeight < cellHeightPx) {
            scale = 1.0;
        }

// Вычисляем новые размеры изображения с учетом масштаба
        int newWidth = (int) (originalWidth * scale);
        int newHeight = (int) (originalHeight * scale);

// Вычисляем отступы по X и Y для центрирования
        double offsetX = (cellWidthPx - newWidth) / 2.0;
        double offsetY = (cellHeightPx - newHeight) / 2.0;

// Конвертируем отступы в единицы измерения Excel (EMU)
        int dx1 = (int) (offsetX * 9525);
        int dy1 = (int) (offsetY * 9525);

// Конвертируем в EMU для правильного позиционирования
        if (workbook instanceof XSSFWorkbook) {
            XSSFDrawing drawing = ((XSSFWorkbook) workbook).getSheetAt(0).createDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, 0, 0, startCol, startRow, endCol + 1, endRow + 1);
            Picture picture = drawing.createPicture(anchor, pictureIdx);
            picture.resize(scale);  // Применяем масштаб
        }

    }
}
