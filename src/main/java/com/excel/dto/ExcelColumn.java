package com.excel.dto;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class ExcelColumn {
    private String header;
    private String fieldName;
    private int width;
}
