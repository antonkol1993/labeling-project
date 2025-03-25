package com.anton.objects;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DefaultItem {
    private Integer ItemNo;
    private String imagePath;
    private String alterImagePath;

    private String mainName;
    private String alterNameRus;

    private String size;
    private String marking;
    private String quantityInBox;
    private String order;
}
