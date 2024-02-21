package com.chaolj.core.excelProvider;

import com.chaolj.core.commonUtils.myServer.Interface.IExcelServer;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
@EnableConfigurationProperties(MyExcelProviderProperties.class)
public class MyExcelProviderConfig {
    @Autowired
    ApplicationContext applicationContext;

    @Autowired
    MyExcelProviderProperties myExcelProviderProperties;

    @Bean(name = "myExcelProvider")
    public IExcelServer MyExcelProvider(){
        return new MyExcelProvider(this.applicationContext, this.myExcelProviderProperties);
    }
}
