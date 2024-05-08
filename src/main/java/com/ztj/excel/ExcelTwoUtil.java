package com.ztj.excel;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Excel的工具类
 * @author 29027
 */
public class ExcelTwoUtil<T> {

    /**
     * 工作薄
     */
    private Workbook wb;

    /**
     * 工作表
     */
    private Sheet sheet;

    /**
     * 需要导出的数据
     */
    private List<T> exportList;

    /**
     * 对象的class对象
     */
    private Class<T> clazz;

    /**
     * 被选中需要导出的字段名称
     */
    private Map<String, Object> checkedFieldsName;

    /**
     * 被选中需要导出的字段对象
     */
    private List<Field> checkedFields;

    /**
     * 包含需要字典转换的字段对象
     */
    private List<Field> fieldsContainDict;

    /**
     * 对象中的字典值
     */
    private Map<String, Map<String, String>> dicts;


    private ExcelTwoUtil() {
    }

    public ExcelTwoUtil(Class<T> clazz) {
        this.clazz = clazz;
    }

    /**
     * @param list
     * @param sheetName
     * @param fieldsName
     */
    public void exportExcel(List<T> list, Map<String, Object> fieldsName, String sheetName, String fileName, HttpServletResponse response) {
        // 初始化数据
        init(list, sheetName, fieldsName);

        // 转换字典值
        try {
            convertDict();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

        // sheet第一行加入名称数据
        createTopRow();

        // sheet其他行，添加目标数据
        try {
            createOtherRow();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

        // 导出wb
        try {
            response.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
            response.setHeader("content-disposition",
                    URLEncoder.encode(sheetName + ".xlsx", "utf-8"));
            //     response.setHeader("Content-disposition",String.format("attachment; filename=\"%s\"", fileName+"-"+System.currentTimeMillis() + ".xlsx"));
            response.setContentType("multipart/form-data");
            response.setCharacterEncoding("UTF-8");
            ServletOutputStream outputStream = response.getOutputStream();
            wb.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                wb.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 添加导出数据
     */
    private void createOtherRow() throws IllegalAccessException {
        Map<String, CellStyle> styles = createStyles(wb);
        for (int rowNum = 1; rowNum <= exportList.size(); rowNum++) {
            Row row = sheet.createRow(rowNum + 1);

//            row.setHeight((short) 600);
            T t = exportList.get(rowNum - 1);
            for (int colNum = 0; colNum < checkedFields.size(); colNum++) {
                Cell cell = row.createCell(colNum);
                cell.setCellStyle(styles.get("data"));
                Field field = checkedFields.get(colNum);
                field.setAccessible(true);
                // 单元格设置值
                addCell(cell, field, t);
            }
        }
    }

    /**
     * 单元格中添加数据
     *
     * @param cell  单元格
     * @param field 字段
     * @param t     list中的一条数据
     */
    private void addCell(Cell cell, Field field, T t) throws IllegalAccessException {
        Class<?> fieldType = field.getType();
        if (String.class == fieldType) {
            if (field.get(t) != null && field.get(t) != "") {
                cell.setCellValue((String) field.get(t));
            }
        } else if ((Integer.TYPE == fieldType) || (Integer.class == fieldType)) {
            if (field.get(t) != null && field.get(t) != "") {
                cell.setCellValue((Integer) field.get(t));
            }
        } else if ((Long.TYPE == fieldType) || (Long.class == fieldType)) {
            if (field.get(t) != null && field.get(t) != "") {
                cell.setCellValue((Long) field.get(t));
            }
        } else if ((Double.TYPE == fieldType) || (Double.class == fieldType)) {
            if (field.get(t) != null && field.get(t) != "") {
                cell.setCellValue((Double) field.get(t));
            }
        } else if ((Float.TYPE == fieldType) || (Float.class == fieldType)) {
            if (field.get(t) != null && field.get(t) != "") {
                cell.setCellValue((Float) field.get(t) == null);
            }
        } else if (Date.class == fieldType) {
            if (field.get(t) != null && field.get(t) != "") {
                String dateFormat = field.getAnnotation(Excel.class).dateFormat();
                cell.setCellValue(dateFormat((Date) field.get(t), dateFormat));
            }
        }
    }

    /**
     * 时间格式转换
     *
     * @param date       日期
     * @param dateFormat 日期格式
     * @return
     */
    private String dateFormat(Date date, String dateFormat) {
        if (dateFormat == null || "".equals(dateFormat)) {
            dateFormat = "yyyy-MM-dd HH:mm:ss";
        }

        SimpleDateFormat df = new SimpleDateFormat(dateFormat);
        return df.format(date);
    }

    /**
     * sheet第一行加入名称数据
     *
     * @param
     */
    private void createTopRow() {
        Row row = this.sheet.createRow(1);
        Map<String, CellStyle> styles = createStyles(wb);
        for (int index = 0; index < checkedFields.size(); index++) {
            Cell cell = row.createCell(index);
            Sheet sheet = cell.getSheet();
            sheet.setColumnWidth(index, 2000);
            row.setHeight((short) 600);
            cell.setCellValue(checkedFields.get(index).getAnnotation(Excel.class).name());
            System.out.println(styles.get("header"));
            cell.setCellStyle(styles.get("header"));
        }
    }


    /**
     * 转换字典值
     * 将数据中字典值转化为对应的值(注:字典值应为String格式)
     */
    private void convertDict() throws IllegalAccessException {
        for (Field field : fieldsContainDict) {
            Excel annotation = field.getAnnotation(Excel.class);
            String dictKey = annotation.dictKey();
            field.setAccessible(true);
            for (T t : exportList) {
                // 获取字段值
                String o = (String) field.get(t);
                field.set(t, dicts.get(dictKey).get(o));
            }
        }
    }

    //获取字段值
    public String mapValue(String dictKey, String childDictKey) {
        //第一层是正常的没有其他符号
        Map<String, String> map = dicts.get(dictKey);
        //第二次因为存在不等于的情况需要获取关系
        String value = dicts.get(dictKey).get(childDictKey);
        if (value == null && ("").equals(value)) {

        }
        return value;
    }

    /**
     * 将数据导出Excel
     *
     * @param list      需要导出的数据
     * @param sheetName 工作表名称
     */
    public void exportExcel(List<T> list, String sheetName, String fileName, HttpServletResponse response) {
        exportExcel(list, null, sheetName, fileName, response);
    }

    /**
     * 将数据导出Excel
     *
     * @param list 需要导出的数据
     */
    /*public void exportExcel(List<T> list) {
        exportExcel(list, null, "sheet");
    }*/

    /**
     * 初始化
     */
    public void init(List<T> list, String sheetName, Map<String, Object> fieldsName) {
        this.checkedFieldsName = fieldsName;

        this.exportList = list;

        // 初始化导出数据
        initExportList();

        // 初始化工作薄
        initWorkbook();

        // 初始化工作表
        initSheet(sheetName, fieldsName);

        // 初始化checkedFields, fieldsContainDict
        initFields();

        // 根据注解生成生成字典
        generateObjDict();
    }

    /**
     * 初始化导出数据
     */
    private void initExportList() {
        // 防止导出过程中出现空指针
        if (Objects.isNull(this.exportList)) {
            this.exportList = new ArrayList<>();
        }
    }

    /**
     * 初始化工作簿
     */
    private void initWorkbook() {
        this.wb = new SXSSFWorkbook();
    }

    /**
     * 初始化工作表
     */
    private void initSheet(String sheetName, Map<String, Object> fieldsName) {
        CellStyle titleStyle = wb.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 18);

        titleStyle.setVerticalAlignment(VerticalAlignment.TOP);
        titleStyle.setFont(titleFont);
        titleFont.setBold(true);
        this.sheet = wb.createSheet();
        Row sheetRow = this.sheet.createRow(0);//表头
        sheetRow.createCell(0).setCellValue(sheetName);
        sheetRow.getCell(0).setCellStyle(titleStyle);
        sheetRow.setHeight((short) 800);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, fieldsName.size() - 1));

    }


    /**
     * 初始化checkedFields, fieldsContainDict
     * fieldsContainDict含有字典表达式的字段
     * checkedFields用户选中的字段
     * 1.如果checkedFieldsName没有定义(未自定义导出字段),所有字段全部导出
     * 2.如果checkedFieldsName进行了定义,根据定义字段进行导出
     */
    private void initFields() {
        // 获取对象所有字段对象
        Field[] fields = clazz.getDeclaredFields();

        // 过滤出checkedFields
        this.checkedFields = Arrays.
                asList(fields).
                stream().
                filter(item -> {
                    if (!Objects.isNull(this.checkedFieldsName)) {
                        if (item.isAnnotationPresent(Excel.class)) {
                            return checkedFieldsName.containsKey(item.getName());
                        }
                    } else {
                        return item.isAnnotationPresent(Excel.class);
                    }
                    return false;
                })
                .collect(Collectors.toList());

        // 过滤出fieldsContainDict
        for (Field declaredField : clazz.getDeclaredFields()) {
            if (declaredField.getAnnotation(Excel.class) != null) {
                System.out.println(declaredField.getAnnotation(Excel.class).dictExp());
            }
        }
        this.fieldsContainDict = Arrays
                .asList(clazz.getDeclaredFields())
                .stream()
                .filter(item -> !"".equals(item.getAnnotation(Excel.class) != null ? item.getAnnotation(Excel.class).dictExp() : ""))
                .collect(Collectors.toList());
    }

    /**
     * 通过扫描字段注解生成字典数据
     */
    private void generateObjDict() {
        if (fieldsContainDict.size() == 0) {
            return;
        }

        if (dicts == null) {
            dicts = new HashMap<>(); //  Map<String, List<Map<String, String>>>
        }

        for (Field field : fieldsContainDict) {
            String dictKey = field.getAnnotation(Excel.class).dictKey();
            String exps = field.getAnnotation(Excel.class).dictExp();
            String[] exp = exps.split(",");

            Map<String, String> keyV = new HashMap<>();

            dicts.put(dictKey, keyV);

            for (String s : exp) {
                String[] out = s.split("=");
                keyV.put(out[0], out[1]);
            }

            System.out.println("字典值:" + dicts);
        }
    }

    /**
     * 创建表格样式
     *
     * @param wb 工作薄对象
     * @return 样式列表
     */
    private Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
        // 数据格式
        CellStyle style = wb.createCellStyle();
        style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        styles.put("data", style);

        // 表头格式
        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        style.setBorderLeft(BorderStyle.DASHED);
        styles.put("header", style);
        return styles;
    }

    public static Date cellDateUtil(String date) {
        try {
            return HSSFDateUtil.getJavaDate(Double.valueOf(Integer.valueOf(date)));
        } catch (Exception exception) {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
            try {
                if (date != null && !("").equals(date)) {
                    return simpleDateFormat.parse(date);
                } else {
                    return null;
                }

            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return null;
    }
}
