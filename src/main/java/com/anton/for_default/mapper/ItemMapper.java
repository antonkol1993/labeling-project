package com.anton.for_default.mapper;



import com.anton.for_default.obj.DefaultItem;
import com.anton.for_default.obj.LabelLargeBox;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class ItemMapper {

    private final Properties mappingToJavaElements = new Properties();
    private final Properties mappingToValueRUS = new Properties();
    private final Properties mappingToImages = new Properties();

    public ItemMapper() {
        loadProperties(mappingToJavaElements, "mappingToJavaElements.properties");
        loadProperties(mappingToValueRUS, "mappingToValueRUS.properties");
        loadProperties(mappingToImages, "mappingToImages.properties");
    }

    private void loadProperties(Properties properties, String fileName) {
        try (InputStream input = getClass().getClassLoader().getResourceAsStream(fileName)) {
            if (input == null) {
                System.err.println("Файл " + fileName + " не найден");
                return;
            }
            properties.load(input);
        } catch (IOException e) {
            System.err.println("Ошибка загрузки файла " + fileName + ": " + e.getMessage());
        }
    }

    public LabelLargeBox map(DefaultItem item) {
        LabelLargeBox labelBox = new LabelLargeBox();
        labelBox.setItemNo(item.getItemNo());
        labelBox.setSize(item.getSize());
        labelBox.setMarking(item.getMarking());
        labelBox.setQuantityInBox(item.getQuantityInBox());
        labelBox.setOrder(item.getOrder());

        // Шаг 1: Найти ключ в mappingToJavaElements по originalName
        String mappedKey = mappingToJavaElements.getProperty(item.getOriginalName());
        if (mappedKey == null) {
            System.err.println("Не найдено соответствие в mappingToJavaElements для: " + item.getOriginalName());
            return labelBox;
        }

        // Шаг 2: Найти значение в mappingToValueRUS по ключу
        String nameRus = mappingToValueRUS.getProperty(mappedKey, "");
        labelBox.setNameRus(nameRus);

        // Шаг 3: Найти значение в mappingToImages по ключу
        String imagePath = mappingToImages.getProperty(mappedKey, "");
        labelBox.setImagePath(imagePath);

        return labelBox;
    }

}
