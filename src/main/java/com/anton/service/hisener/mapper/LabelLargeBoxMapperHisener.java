package com.anton.service.hisener.mapper;

import com.anton.objects.InvoiceItemData;
import com.anton.objects.LabelLargeBox;
import com.anton.service.hisener.reader.ExcelInvoiceParser;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.stream.Collectors;

public class LabelLargeBoxMapperHisener {
    public static void main(String[] args) throws IOException {
//        String filePath = "excel-example/China14 invoices/25HS10047P-PI  Final 3.13.xlsx";
        String filePath = "25HS10038P-PI.xlsx";

        List<List<InvoiceItemData>> parsedData = ExcelInvoiceParser.parseExcel(filePath);

        LabelLargeBoxMapperHisener mapper = new LabelLargeBoxMapperHisener();
        List<List<LabelLargeBox>> mappedData = mapper.mapData(parsedData);

        // Выводим результат
        mappedData.forEach(group -> {
            System.out.println("=== Новая группа ===");
            group.forEach(item -> System.out.println(
                    String.join(" | ",
                            item.getNameRus(),
                            item.getSize(),
                            item.getMarking(),
                            item.getQuantityInBox(),
                            item.getOrder(),
                            item.getImagePath()
                    )
            ));
        });
    }



    private final Map<String, String> invoiceMapping;
    private final Map<String, String> rusMapping;
    private final Map<String, String> imageMapping;

    public LabelLargeBoxMapperHisener() throws IOException {
        this.invoiceMapping = loadProperties("mapping_item-invoice.properties");
        this.rusMapping = loadProperties("mapping_item-RUSvalue.properties");
        this.imageMapping = loadProperties("mapping_item-image.properties");
    }

    private Map<String, String> loadProperties(String fileName) throws IOException {
        Properties properties = new Properties();
        try (InputStream input = getClass().getClassLoader().getResourceAsStream(fileName);
             Reader reader = new InputStreamReader(input, StandardCharsets.UTF_8)) {
            if (input == null) {
                throw new IOException("Файл " + fileName + " не найден.");
            }
            properties.load(reader);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return properties.entrySet().stream()
                .collect(Collectors.toMap(e -> e.getKey().toString(), e -> e.getValue().toString()));
    }

    public List<List<LabelLargeBox>> mapData(List<List<InvoiceItemData>> invoiceData) {
        return invoiceData.stream()
                .map(group -> group.stream()
                        .map(this::mapItem)
                        .collect(Collectors.toList()))
                .collect(Collectors.toList());
    }

    private LabelLargeBox mapItem(InvoiceItemData item) {
        LabelLargeBox label = new LabelLargeBox();

        // Ищем ключ, у которого значение равно item.getName()
        String mappedKey = invoiceMapping.entrySet().stream()
                .filter(entry -> entry.getValue().equals(item.getName()))
                .map(Map.Entry::getKey)
                .findFirst()
                .orElse(item.getName()); // Если не нашли, оставляем оригинальное имя

        // Теперь ищем русское название по найденному ключу
        label.setNameRus(rusMapping.getOrDefault(mappedKey, "Неизвестно"));
        label.setImagePath(imageMapping.getOrDefault(mappedKey, "Нет изображения"));
        label.setInvoiceItemNumber(item.getElementNumber());
        label.setSize(item.getSize());
        label.setMarking(item.getFinish());
        label.setQuantityInBox(String.valueOf(item.getBox()));
        label.setOrder(item.getProformaNo());
        label.setNameAndSize(label.getNameRus() + "\n" + label.getSize());

        return label;
    }
}

