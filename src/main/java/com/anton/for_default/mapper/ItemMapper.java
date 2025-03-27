package com.anton.for_default.mapper;

import com.anton.for_default.obj.DefaultItem;
import com.anton.for_default.obj.LabelLargeBox;
import com.anton.for_default.reader.DefaultReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class ItemMapper {


    public static void main(String[] args) {
        Logger logger = LoggerFactory.getLogger(ItemMapper.class);
        String filePath = "excel-example/DataFromInvoice for example .xlsx"; // Укажите путь к файлу

        DefaultReader reader = new DefaultReader();
        ItemMapper mapper = new ItemMapper();

        try {
            // Шаг 1: Читаем Excel
            List<List<DefaultItem>> dataBlocks = reader.readExcel(filePath);
            logger.info("Прочитано {} блоков данных.", dataBlocks.size());

            // Шаг 2: Преобразуем в LabelLargeBox
            List<List<LabelLargeBox>> mappedBlocks = mapper.map(dataBlocks);
            logger.info("Маппинг завершён. Получено {} блоков LabelLargeBox.", mappedBlocks.size());

            // Шаг 3: Вывод результатов
            for (int i = 0; i < mappedBlocks.size(); i++) {
                System.out.printf("Блок #%d%n", i + 1);
                for (LabelLargeBox label : mappedBlocks.get(i)) {
                    System.out.println(label);
                }
                System.out.println("----------------------");
            }

        } catch (IOException e) {
            logger.error("Ошибка при обработке файла: {}", e.getMessage());
        }
    }

    private final Properties mappingToImage = new Properties();
    private final Properties mappingToValueRUS = new Properties();
    private final Properties mappingToImages = new Properties();

    public ItemMapper() {
        loadProperties(mappingToImage, "mapping_item-invoice.properties");
        loadProperties(mappingToValueRUS, "mapping_item-RUSvalue.properties");
        loadProperties(mappingToImages, "mapping_item-image.properties");
    }

    private void loadProperties(Properties properties, String fileName) {
        try (InputStream input = getClass().getClassLoader().getResourceAsStream(fileName);
             InputStreamReader reader = new InputStreamReader(input, StandardCharsets.UTF_8)) {

            if (input == null) {
                System.err.println("Файл " + fileName + " не найден");
                return;
            }

            properties.load(reader);
        } catch (IOException e) {
            System.err.println("Ошибка загрузки файла " + fileName + ": " + e.getMessage());
        }
    }


    public List<List<LabelLargeBox>> map(List<List<DefaultItem>> dataBlocks) {
        List<List<LabelLargeBox>> mappedBlocks = new ArrayList<>();

        for (List<DefaultItem> block : dataBlocks) {
            List<LabelLargeBox> mappedBlock = new ArrayList<>();
            for (DefaultItem item : block) {
                mappedBlock.add(mapItem(item));
            }
            mappedBlocks.add(mappedBlock);
        }

        return mappedBlocks;
    }

    private LabelLargeBox mapItem(DefaultItem item) {
        LabelLargeBox labelBox = new LabelLargeBox();
        labelBox.setItemNo(item.getItemNo());
        labelBox.setSize(item.getSize());
        labelBox.setMarking(item.getMarking());
        labelBox.setQuantityInBox(item.getQuantityInBox());
        labelBox.setOrder(item.getOrder());

        // 📌 Очищаем originalName от лишних пробелов
        String originalName = item.getOriginalName();
        if (originalName == null || originalName.trim().isEmpty()) {

            System.err.println("⚠ Ошибка: originalName == null или пустой!" + "No " + '[' + item.getItemNo() + ']');
        }
        originalName = originalName.trim();

        // 🔍 **Шаг 1: Ищем originalName среди значений в mapping_item-invoice.properties**
        String mappedKey = findKeyByValue(mappingToImage, originalName);
        if (mappedKey == null) {
            System.err.println("❌ Значение [" + originalName + "] не найдено в mapping_item-invoice.properties!");
            labelBox.setNameRus(item.getAlterNameRus());
            labelBox.setImagePath(item.getAlterImagePath());
            return labelBox;
        } else {
            System.out.println("✅ Найден ключ: " + mappedKey);
            labelBox.setKeyName(mappedKey);
            // 🔍 **Шаг 2: Находим перевод в mapping_item-RUSvalue.properties**
            String nameRus = mappingToValueRUS.getProperty(mappedKey, "");
            labelBox.setNameRus(nameRus);

            // 🔍 **Шаг 3: Находим путь к изображению в mapping_item-image.properties**
            String imagePath = mappingToImages.getProperty(mappedKey, "");
            labelBox.setImagePath(imagePath);
            return labelBox;
        }

    }

    private String findKeyByValue(Properties properties, String valueToFind) {
        for (String key : properties.stringPropertyNames()) {
            if (properties.getProperty(key).trim().equals(valueToFind)) {
                return key; // Возвращаем ключ, если нашли совпадение по значению
            }
        }
        return null; // Если не нашли, возвращаем null
    }


}
