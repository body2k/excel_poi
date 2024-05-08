package com.ztj.excel;

import com.alibaba.excel.util.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Excel的工具类
 */
public class ExcelUtil<T> {

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


    private ExcelUtil() {
    }

    public ExcelUtil(Class<T> clazz) {
        this.clazz = clazz;
    }

    /**
     * @param list
     * @param sheetName
     * @param fieldsName
     */
    public void exportExcel(List<T> list, Map<String, Object> fieldsName, String sheetName,
                            String fileName,
                            HttpServletResponse response) {
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
            row.setHeight((short) 600);
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
            sheet.setColumnWidth(index, 5000);
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
        titleFont.setFontHeightInPoints((short) 15);

        titleStyle.setFont(titleFont);
        this.sheet = wb.createSheet();
        Row sheetRow = this.sheet.createRow(0);//表头
        sheetRow.createCell(0).setCellValue(sheetName);
        sheetRow.getCell(0).setCellStyle(titleStyle);
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

    /**
     * 设置单元
     */
    public static void setCell(Object object, Cell cell, Field[] fields, Integer i) throws IllegalAccessException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        for (Field field : fields) {
            field.setAccessible(true);

            if (cell.getStringCellValue().contains("{" + field.getName() + "}")) {
                cell.setCellType(CellType.STRING);
                String stringCellValue = cell.getStringCellValue();

                String s = stringCellValue.replaceAll("\\{" + field.getName() + "}", field.get(object) == null ? "" : field.getGenericType().toString().equals("class java.util.Date") ? simpleDateFormat.format(field.get(object)) : field.get(object).toString());
                cell.setCellValue(s);
            }
        }
    }


    /**
     * 默认读取全部的数据
     */
    public static List<List<Object>> readALLExcel(XSSFWorkbook xssfWorkbook, Integer sheetAt, Integer skip) {
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(sheetAt);
        List<List<Object>> list = new LinkedList<>();
        for (Row cells : xssfSheet) {
            if (!isRowEmpty(cells)) {
                List<Object> rowList = new LinkedList<>();
                for (int i = 0; i < xssfSheet.getRow(skip).getLastCellNum(); i++) {
                    Cell cell = cells.getCell(i);
                    rowList.add(ExcelUtil.getCellStringVal(cell));
                }
//                for (Cell cell : cells) {
//
////                cell.setCellValue();
//                    rowList.add(ExcelUtil.getCellStringVal(cell));
//                }
                list.add(rowList);
            }
        }
        return list.stream().skip(skip).collect(Collectors.toList());

    }


    public static List<List<Object>> readALLExcel(XSSFSheet sheetAt, Integer headerIndex) {
        List<List<Object>> list = new LinkedList<>();
        short lastCellNum = sheetAt.getRow(headerIndex - 1).getLastCellNum();
        for (Row cells : sheetAt) {
            if (!isRowEmpty(cells)) {
                List<Object> rowList = new LinkedList<>();
                if (cells.getRowNum() >= headerIndex) {
                    for (int i = 0; i < lastCellNum; i++) {
                        Cell cell = cells.getCell(i);
                        rowList.add(ExcelUtil.getCellStringVal(cell));

                    }
//                    for (Cell cell : cells) {
//                    }
                    list.add(rowList);
                }
            }
        }
        return list;

    }

    public static List<Object> readRowExcel(XSSFSheet sheetAt, Integer headerIndex) {
        List<Object> list = new LinkedList<>();
        XSSFRow row = sheetAt.getRow(headerIndex);
        for (Cell cell : row) {
            list.add(ExcelUtil.getCellStringVal(cell));
        }
        return list;
    }

    public static String getCellStringVal(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellTypeEnum()) {
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
                    Date javaDate = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                    return simpleDateFormat.format(javaDate);
                } else {
                    cell.setCellType(CellType.STRING);
                    return String.valueOf(cell.getStringCellValue());

                }
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            case ERROR:
                return String.valueOf(cell.getErrorCellValue());
            default:
                return "";
        }
    }


    public static <T> List<T> readList(XSSFWorkbook xssfWorkbook, T t, Integer sheetAt, Integer headerIndex) throws IllegalAccessException {
        List<T> list = new LinkedList<>();
        //存储
        Map<Integer, String> map = new HashMap<>();
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(sheetAt);
        Field[] declaredFields = t.getClass().getDeclaredFields();
        XSSFRow row = xssfSheet.getRow(headerIndex);
        short lastCellNum = xssfSheet.getRow(headerIndex).getLastCellNum();
        //指定行读出来的一定是string类型
        for (Field field : declaredFields) {
            for (Cell cell : row) {
                String tableName = cell.getStringCellValue();
                String name = field.getAnnotation(Excel.class).name();
                if (name.equals(tableName)) {
                    String value = field.getName();

                    /**
                     * 存储标和表格对应的字段每次
                     * */
                    map.put(cell.getColumnIndex(), value);
                }

            }

        }
        //读内容Type parameter 'T' cannot be instantiated directly
        T t1;
        for (Row rows : xssfSheet) {
            if (!isRowEmpty(rows)) {

                if (rows.getRowNum() > headerIndex) {
                    t1 = t;
                    for (int i = 0; i < lastCellNum; i++) {
                        Cell cell = rows.getCell(i);
                        int columnIndex = cell.getColumnIndex();
                        //获取字段每次
                        String valueName = map.get(columnIndex);
                        Field[] declaredFields1 = t1.getClass().getDeclaredFields();
                        for (Field field : declaredFields1) {
                            if (field.getName().equals(valueName)) {
                                field.set(t1, ExcelUtil.getCellStringVal(cell));
                            }
                        }
                    }


                    list.add(t);
                }
            }
        }
        return list;
    }

    public static void setCell(String value, Integer integer, Row row) {
        Cell cell = row.createCell(integer);
        cell.setCellValue(value);
    }

    //判断row是否为空
    public static boolean isRowEmpty(Row row) {
        if (null == row) {
            return true;
        }
        int firstCellNum = row.getFirstCellNum();   //第一个列位置
        int lastCellNum = row.getLastCellNum();     //最后一列位置
        int nullCellNum = 0;    //空列数量
        for (int c = firstCellNum; c < lastCellNum; c++) {
            Cell cell = row.getCell(c);
            if (null == cell || Cell.CELL_TYPE_BLANK == cell.getCellType()) {
                nullCellNum++;
                continue;
            }
            cell.setCellType(Cell.CELL_TYPE_STRING);
            String cellValue = cell.getStringCellValue().trim();
            if (StringUtils.isEmpty(cellValue)) {
                nullCellNum++;
            }
        }
        //所有列都为空
        if (nullCellNum == (lastCellNum - firstCellNum)) {
            return true;
        }
        return false;
    }


    /**
     * 创建模板设置
     */
    public static void createExcel(XSSFWorkbook workbook, List<List<Object>> list, Integer index) {
        XSSFSheet sheetAt = workbook.getSheetAt(index);


        //循环行
        for (int rowIndex = 0; rowIndex < list.size(); rowIndex++) {

            XSSFRow row = sheetAt.createRow(rowIndex);
            row.setHeight((short) 600);
            //循环列
            List<Object> listCell = list.get(rowIndex);
            for (int cellIndex = 0; cellIndex < listCell.size(); cellIndex++) {
                XSSFCell cell = row.createCell(cellIndex);
                sheetAt.setColumnWidth(cellIndex, 5000);
                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
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
                font.setFontHeightInPoints((short) 15);
                style.setFont(font);
                cell.setCellStyle(style);
                cell.setCellValue(listCell.get(cellIndex).toString());
            }
        }
        mergeRows(workbook, index);
//        setTheStyle(workbook,index);
    }

    /**
     * 合并行
     */
    public static void mergeRows(XSSFWorkbook workbook, Integer index) {
        XSSFSheet sheetAt = workbook.getSheetAt(index);
        String value = "";
        Integer lastKey = null;
        Integer lastValue = null;
        List<LinkedHashMap<Integer, Integer>> mergeList = new LinkedList<LinkedHashMap<Integer, Integer>>();
        for (Row row : sheetAt) {
            LinkedHashMap<Integer, Integer> initMap = new LinkedHashMap<>();
            if (!isRowEmpty(row)) {
                for (Cell cell : row) {
                    String stringCellValue = getCellStringVal(cell);
                    //如果相同则合并不相同则不合并
                    if (stringCellValue.equals(value)) {
                        for (Map.Entry<Integer, Integer> entry : initMap.entrySet()) {
                            lastKey = entry.getKey();
                        }
                        initMap.remove(lastKey);
                        initMap.put(lastKey, cell.getColumnIndex());
                    } else {
                        initMap.put(cell.getColumnIndex(), cell.getColumnIndex());
                        value = stringCellValue;
                    }

                }
            }
            mergeList.add(initMap);
        }
        for (int i = 0; i < mergeList.size(); i++) {
            LinkedHashMap<Integer, Integer> map = mergeList.get(i);
            for (Map.Entry<Integer, Integer> entry : map.entrySet()) {
                Integer key = entry.getKey();
                Integer value1 = entry.getValue();
                //结束行大于合并行则合并
                if (value1 - key > 0) {

                    sheetAt.addMergedRegion(new CellRangeAddress(i, i, key, value1));
                }

            }
        }
    }


    //Excel单元格插入图片
    public static void cellImage(Integer row,Integer col,String base64,Sheet sheet,Workbook wb) throws Exception {
        Drawing patriarch = sheet.createDrawingPatriarch();
        ClientAnchor anchor = new XSSFClientAnchor(0, 0, 255, 255, col,row,col+1,row+1);
        BASE64Decoder decoder = new BASE64Decoder();

        byte[] base64Bytes = decoder.decodeBuffer(base64);
        System.out.println(base64Bytes);
        if (base64!=null&&!("").equals(base64)){
            patriarch.createPicture(anchor, wb.addPicture(base64Bytes, HSSFWorkbook.PICTURE_TYPE_JPEG));

        }
    }
}


//    /**
//     * 设置样式
//     * */
//    public static void setTheStyle(XSSFWorkbook workbook,Integer index){
//        XSSFSheet sheetAt = workbook.getSheetAt(index);
//        for (Row row:sheetAt){
//            for (Cell cell:row){
//
//            }
//        }
//    }

