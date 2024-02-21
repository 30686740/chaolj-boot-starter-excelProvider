package com.chaolj.core.excelProvider.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class CellValue {
    // 字符串值
    private String strValue;
    // 字段类型
    private String fieldType;
    // 字段类型
    private String format;
}
