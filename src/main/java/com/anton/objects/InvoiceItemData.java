package com.anton.objects;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class InvoiceItemData {
    private String proformaNo;
    private String name;
    private Integer elementNumber;

    private String size;
    private String partNo;
    private String finish;
    private Integer box;
    private String ctn;
    private String ctnPltCtn;
    private Double kgsUn;
    private Double totalKgs;
    private Double quantity;
    private String unit;
    private Double unitPrice;
    private Double total;

}
