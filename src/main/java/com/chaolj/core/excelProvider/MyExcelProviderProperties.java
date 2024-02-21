package com.chaolj.core.excelProvider;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;

@Data
@ConfigurationProperties(prefix = "myproviders.myexcelprovider")
public class MyExcelProviderProperties {
    private String defaultFileName = "工作簿";
    private String defaultSheetName = "Sheet1";
}
