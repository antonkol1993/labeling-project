package com.anton.service.generator;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;

public class ImageHandlerNew {

    public static void addImageToSheet(Workbook workbook, Sheet sheet, String imagePath,
                                       int startRow, int startCol, int endRow, int endCol) throws IOException {
        // Загружаем изображение в массив байтов
        byte[] imageBytes;
        try (InputStream inputStream = new FileInputStream(imagePath)) {
            imageBytes = IOUtils.toByteArray(inputStream);
        } catch (IOException e) {
            System.err.println("Image not found: " + imagePath);
            e.printStackTrace();
            throw e;
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

        // --- МАСШТАБИРОВАНИЕ ---
        double scale = 1.0;  // По умолчанию — не изменяем размер

        // Если изображение больше ячейки, то уменьшаем его
        if (originalWidth > cellWidthPx || originalHeight > cellHeightPx) {
            double scaleX = (cellWidthPx * 0.80) / originalWidth;  // 80% от ширины ячейки
            double scaleY = (cellHeightPx * 0.80) / originalHeight; // 80% от высоты ячейки
            scale = Math.min(scaleX, scaleY);
        }

        // Если изображение меньше или равно ячейке, **оставляем его как есть**
        int newWidth = (int) (originalWidth * scale);
        int newHeight = (int) (originalHeight * scale);

        // Вычисляем отступы для центрирования
        double offsetX = (cellWidthPx - newWidth) / 2.0;
        double offsetY = (cellHeightPx - newHeight) / 2.0;

        // Конвертируем отступы в EMU (единицы Excel)
        int dx1 = (int) (offsetX * 9525);
        int dy1 = (int) (offsetY * 9525);

        // Добавляем изображение в Excel
        if (workbook instanceof XSSFWorkbook) {
            XSSFDrawing drawing = ((XSSFWorkbook) workbook).getSheetAt(0).createDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, -dx1, -dy1, startCol, startRow, endCol + 1, endRow + 1);
//            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
            Picture picture = drawing.createPicture(anchor, pictureIdx);
        }
    }
}
