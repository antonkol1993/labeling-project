package com.anton.labeling.objects;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ItemLargeBox {
    private String name;
    private String size;
    private String marking;
    private String quantityInBox;
    private String order;

    private String nameAndSize;
    private Integer invoiceItemNumber;
}
