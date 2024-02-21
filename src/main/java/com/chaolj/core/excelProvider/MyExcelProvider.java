package com.chaolj.core.excelProvider;

import cn.hutool.core.convert.Convert;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
import com.chaolj.core.excelProvider.entity.CellValue;
import com.chaolj.core.commonUtils.myServer.Interface.IExcelServer;
import com.chaolj.core.excelProvider.annotation.ExcelField;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.context.ApplicationContext;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.*;
import javax.servlet.http.HttpServletResponse;

@Slf4j
public class MyExcelProvider implements IExcelServer {
    public static final String EXCEL_XLS = ".xls";
    public static final String EXCEL_XLSX = ".xlsx";
    public static final String[] NUM_TYPE = {"int", "long", "integer", "double", "BigDecimal", "Short"};

    private ApplicationContext applicationContext;
    private MyExcelProviderProperties properties;

    public MyExcelProvider(ApplicationContext applicationContext, MyExcelProviderProperties properties) {
        this.applicationContext = applicationContext;
        this.properties = properties;
    }

    @Override
    public <T> byte[] transfer(List<T> data, Class<T> clazz, String sheetName) {
        if (StrUtil.isBlank(sheetName)) sheetName = properties.getDefaultSheetName();

        byte[] bytes = null;

        try {
            // 创建Workbook
            var sxwb = new SXSSFWorkbook(300);
            // 创建Sheet
            var sheet = sxwb.createSheet(sheetName);
            // 填充Sheet
            this.fillSheet(sxwb, sheet, data, clazz);

            // 将文件输出
            var ouputStream = new ByteArrayOutputStream();
            sxwb.write(ouputStream);

            bytes = ouputStream.toByteArray();
        }
        catch (Exception ex) {
            log.error("MyExcelProvider.transfer 发生错误！", ex);
        }

        return bytes;
    }

    @Override
    public <T> void download(HttpServletResponse response, List<T> data, Class<T> clazz, String fileName, String sheetName) {
        if (StrUtil.isBlank(fileName)) fileName = properties.getDefaultFileName();
        if (StrUtil.isBlank(sheetName)) sheetName = properties.getDefaultSheetName();

        // 设置 http response 响应头
        this.setResponseHeader(response, fileName);

        try {
            // 创建Workbook
            var sxwb = new SXSSFWorkbook(300);
            // 创建Sheet
            var sheet = sxwb.createSheet(sheetName);
            // 填充Sheet
            this.fillSheet(sxwb, sheet, data, clazz);

            //将文件输出
            var ouputStream = response.getOutputStream();
            sxwb.write(ouputStream);

            ouputStream.flush();
            ouputStream.close();
        }
        catch (Exception ex){
            log.error("MyExcelProvider.download 发生错误！", ex);
        }
    }

    private <T> void fillSheet(SXSSFWorkbook sxwb, Sheet sheet, List<T> dataList, Class<T> clazz) {

        // region 填充表头

        Row row = null;
        Cell cell = null;

        // 第一行，列头
        row = sheet.createRow(0);

        var headList = this.getExcelHeadFromClass(clazz);
        if (headList != null && !headList.isEmpty()) {
            for (int i = 0; i < headList.size(); i++) {
                cell = row.createCell(i);
                cell.setCellValue(headList.get(i).get("headName"));
                sheet.setColumnWidth(i, Integer.parseInt(headList.get(i).get("width")) * 256 + 184);
            }
        }

        // endregion

        // region 填充内容

        if (dataList == null || dataList.isEmpty()) return;

        var numberStyle = sxwb.createCellStyle();
        numberStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("0.00"));

        var cellValue = new CellValue();
        for (int rowIndex = 0; rowIndex < dataList.size(); rowIndex++) {
            row = sheet.createRow(rowIndex+1);

            for (int cellIndex = 0; cellIndex < headList.size(); cellIndex++) {
                cell = row.createCell(cellIndex);
                cellValue = this.getValueOfClass(headList.get(cellIndex), dataList.get(rowIndex), clazz);

                if (StrUtil.containsAnyIgnoreCase(cellValue.getFieldType(), NUM_TYPE)) {
                    if (StrUtil.isNotBlank(cellValue.getStrValue()) && NumberUtil.isNumber(cellValue.getStrValue())) {
                        cell.setCellValue(Double.parseDouble(cellValue.getStrValue()));
                        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        if (StrUtil.isNotBlank(cellValue.getFormat())) {
                            numberStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat(cellValue.getFormat()));
                            cell.setCellStyle(numberStyle);
                        }
                    } else {
                        cell.setCellValue(cellValue.getStrValue());
                    }
                } else {
                    cell.setCellValue(cellValue.getStrValue());
                }
            }
        }

        // endregion

        sheet.createFreezePane(0,1);
    }

    private List<Map<String, String>> getExcelHeadFromClass(Class clazz) {
        List<Map<String, String>> headList = new ArrayList<>(16);

        //循环注解里面的值
        Field[] fieldList = clazz.getDeclaredFields();

        Arrays.stream(fieldList).filter(field -> {
            ExcelField excelFieldAnno = field.getDeclaredAnnotation(ExcelField.class);
            return excelFieldAnno != null;
        }).sorted(Comparator.comparing(field -> {
            ExcelField excelFieldAnno = field.getDeclaredAnnotation(ExcelField.class);
            return excelFieldAnno.order();
        })).forEach(field -> {
            ExcelField excelFieldAnno = field.getDeclaredAnnotation(ExcelField.class);
            Map<String, String> annoMap = new HashMap<>(16);
            annoMap.put("fieldName", field.getName());
            annoMap.put("headName", StrUtil.isNotBlank(excelFieldAnno.name()) ? excelFieldAnno.name() : field.getName());
            annoMap.put("order", String.valueOf(excelFieldAnno.order()));
            annoMap.put("dateFormat", excelFieldAnno.dateFormat());
            annoMap.put("numFormat", excelFieldAnno.numFormat());
            annoMap.put("suffix", excelFieldAnno.suffix());
            annoMap.put("width", String.valueOf(excelFieldAnno.width()));
            annoMap.put("type", String.valueOf(excelFieldAnno.type()));
            annoMap.put("divide",excelFieldAnno.divide());
            annoMap.put("place",String.valueOf(excelFieldAnno.place()));
            headList.add(annoMap);
        });

        return headList;
    }

    private CellValue getValueOfClass(Map<String, String> headMap, Object obj, Class clazz) {
        String value = "";
        try {
            Field field = clazz.getDeclaredField(headMap.get("fieldName"));
            field.setAccessible(true);

            // 属性类型
            String fieldType = field.getGenericType().toString();
            Method method = ReflectUtil.getMethod(clazz, "get"+this.UpFirstCase(headMap.get("fieldName")));
            Object fieldValue = method.invoke(obj);

            if (fieldType.contains("String")) {
                value = (null == fieldValue) ? "" : fieldValue.toString();
            } else if (fieldType.contains("Date")) {
                Date date = (null == fieldValue) ? null : (Date) fieldValue;
                if (date == null) {
                    value = "";
                } else {
                    value = DateUtil.format(date, headMap.get("dateFormat"));
                }
            } else if (StrUtil.containsAnyIgnoreCase(fieldType, NUM_TYPE)) {
                BigDecimal num = Convert.toBigDecimal(fieldValue, BigDecimal.ZERO);

                String format = headMap.get("numFormat");
                String divide = headMap.get("divide");
                if (StrUtil.isNotBlank(divide) || StrUtil.isNotBlank(format)) {
                    int type = Integer.parseInt(headMap.get("type"));
                    // 是否触发除法计算
                    if (StrUtil.isBlank(divide)) {
                        DecimalFormat decimalFormat = new DecimalFormat(headMap.get("numFormat"));
                        decimalFormat.setRoundingMode(RoundingMode.valueOf(type));
                        value = decimalFormat.format(num);
                    } else {
                        int place = Integer.parseInt(headMap.get("place"));
                        value = num.divide(new BigDecimal(divide), place, RoundingMode.valueOf(type)).toString();
                    }
                } else {
                    value = num.toPlainString();
                }
            }
            return new CellValue(value + headMap.get("suffix"), fieldType, headMap.get("numFormat"));
        } catch (NoSuchFieldException | IllegalAccessException | InvocationTargetException noSuchFieldException) {
            log.error(">>>>>>>>>> PoiUtil.getValueOfClass() ", noSuchFieldException);
        }
        return new CellValue("","", headMap.get("numFormat"));
    }

    private String UpFirstCase(String str) {
        char[] chars = str.toCharArray();
        if(!Character.isUpperCase(chars[0])){
            return str.substring(0,1).toUpperCase()+str.substring(1);
        }
        return String.valueOf(chars);

    }

    private void setResponseHeader(HttpServletResponse response, String filename) {
        String fileName = filename + DateUtil.format(new Date(), "yyyy-MM-dd-HHmmss") + EXCEL_XLSX;
        try {
            response.reset();
            response.setContentType("application/vnd.ms-excel;charset=UTF-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + new String(fileName.getBytes(), "ISO8859-1"));
            response.addHeader("Pargam", "no-cache");
            response.addHeader("Cache-Control", "no-cache");
            response.addHeader("Access-Control-Expose-Headers", "filename,Content-Disposition");
            response.addHeader("filename", fileName);
        } catch (Exception ex) {
            log.error("MyExcelProvider.setResponseHeader 发生错误！", ex);
        }
    }
}
