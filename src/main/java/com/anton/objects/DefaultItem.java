package com.anton.objects;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DefaultItem {
    private Integer itemNo;
    private String imagePath;
    private String alterImagePath;

    private String originalName;
    private String alterNameRus;

    private String size;
    private String marking;
    private String quantityInBox;
    private String order;
    @Override
    public String toString() {
        return  '{' + "itemNo=" + itemNo +
                " imagePath=" + imagePath + " | " +
                " alterImagePath=" + alterImagePath + " | " +
                " originalName=" + originalName + " | " +
                " alterNameRus=" + alterNameRus + " | " +
                " size=" + size + " | " +
                " marking=" + marking + " | " +
                " quantityInBox=" + quantityInBox + " | " +
                " order=" + order + '}' ;
    }

}
