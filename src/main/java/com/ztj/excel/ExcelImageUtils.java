package com.ztj.excel;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.util.*;

/*** 解析excel内的图片所在单元格 利用Poi的方法
 * @author 29027
 */
public class ExcelImageUtils {

    private final static Logger LOG = LoggerFactory.getLogger(ExcelImageUtils.class);

    /**
     * 根据文件解析excel内图片
     *
     * @param input      文件对象
     * @param sheetIndex 第几个sheet
     * @return
     */
    public static Map<String, Map<Integer, PictureData>> parseSheetPic(InputStream input, int sheetIndex) {
        try {
            System.out.println(input);
            System.out.println("----------------");
            //实现并返回

            return parseSheetPics(input, sheetIndex);
        } catch (Exception e) {
            LOG.error("parseSheetPic fail , file error={}", e);
        }
        //默认
        return new HashMap<>();
    }

    /**
     * 根据文件流及文件类型解析excel内图片
     *
     * @param inputStream 文件流
     * @param sheetIndex  第几个sheet
     * @return
     */
    public static Map<String, Map<Integer, PictureData>> parseSheetPics(InputStream inputStream, int sheetIndex) {
        try {
            //初始化
            Workbook workbook = WorkbookFactory.create(inputStream);
            //获取sheet
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            //第一行作为key
            Map<Integer, String> keyMap = new HashMap<>();
            //第一行的迭代器
            Iterator<Cell> iterator = sheet.getRow(0).cellIterator();
            //指针
            int p = 0;
            //如果存在
            while (iterator.hasNext()) {
                //获取值
                String value = iterator.next().getStringCellValue();
                //如果存在
                if (value != null && !("").equals(value)) {
                    //记录该key
                    keyMap.put(p, value);
                }
                //无论如何+1o
                p++;
            }

            return parseSheetPicForXLSX((XSSFWorkbook) workbook, sheetIndex, keyMap);

        } catch (Exception e) {
            LOG.error("parseSheetPic fail , inputStream error={}", e);
        }
        //默认
        return new HashMap<>();
    }

    /**
     * 根据文件流及文件类型解析excel内图片
     *
     * @param inputStream 文件流
     * @param sheetIndex  第几个sheet
     * @return
     */
    public static Map<String, Map<Integer, PictureData>> parseSheetPics(   Workbook workbook, int sheetIndex) {
        try {

            //获取sheet
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            //第一行作为key
            Map<Integer, String> keyMap = new HashMap<>();
            //第一行的迭代器
            Iterator<Cell> iterator = sheet.getRow(0).cellIterator();
            //指针
            int p = 0;
            //如果存在
            while (iterator.hasNext()) {
                //获取值
                String value = iterator.next().getStringCellValue();
                //如果存在
                if (value != null && !("").equals(value)) {
                    //记录该key
                    keyMap.put(p, value);
                }
                //无论如何+1o
                p++;
            }

            return parseSheetPicForXLSX((XSSFWorkbook) workbook, sheetIndex, keyMap);

        } catch (Exception e) {
            LOG.error("parseSheetPic fail , inputStream error={}", e);
        }
        //默认
        return new HashMap<>();
    }

/**
 * header 第几行是表头
 * */
    public static Map<String, Map<Integer, PictureData>> parseSheetPics(  XSSFSheet sheet,Integer header ) {
        try {


            //第一行作为key
            Map<Integer, String> keyMap = new HashMap<>();
            //第一行的迭代器
            Iterator<Cell> iterator = sheet.getRow(header).cellIterator();
            //指针
            int p = 0;
            //如果存在
            while (iterator.hasNext()) {
                //获取值
                String value = iterator.next().getStringCellValue();
                //如果存在
                if (value != null && !("").equals(value)) {
                    //记录该key
                    keyMap.put(p, value);
                }
                //无论如何+1o
                p++;
            }
            System.out.println("-------------------");
System.out.println(keyMap);
            return parseSheetPicForXLSX(sheet, keyMap);

        } catch (Exception e) {
            LOG.error("parseSheetPic fail , inputStream error={}", e);
        }
        //默认
        return new HashMap<>();
    }


    /**
     * 解析XLS的图片
     *
     * @param workbook   workbook
     * @param sheetIndex 第几页sheet的图片
     * @return
     */
    public static List<PictureData> parseSheetPicForXLS(XSSFSheet sheetAt) {
//        List<HSSFPictureData> pictures = workbook.getAllPictures();
        List<POIXMLDocumentPart> relations = sheetAt.getRelations();
        List<PictureData> list = new LinkedList<>();
        for (POIXMLDocumentPart poixmlDocumentPart : relations) {
            if (poixmlDocumentPart instanceof XSSFDrawing) {
                //强转
                XSSFDrawing drawing = (XSSFDrawing) poixmlDocumentPart;
                //获取列表
                List<XSSFShape> shapeList = drawing.getShapes();
                //循环2
                for (XSSFShape shape : shapeList) {
                    //强转
                    XSSFPicture pic = (XSSFPicture) shape;
                    list.add(pic.getPictureData());
                }
            }
        }
        return list;
    }

    /**
     * 解析XLSX图片
     *
     * @return
     */
    private static Map<String, Map<Integer, PictureData>> parseSheetPicForXLSX
    (XSSFWorkbook workbook, int sheetIndex, Map<Integer, String> keyMap) {
        //初始化结果
        Map<String, Map<Integer, PictureData>> result = new HashMap<>();
        //获取列表
        List<POIXMLDocumentPart> relationList = workbook.getSheetAt(sheetIndex).getRelations();
        //循环1
        for (POIXMLDocumentPart dr : relationList) {
            //如果是类型
            if (dr instanceof XSSFDrawing) {
                //强转
                XSSFDrawing drawing = (XSSFDrawing) dr;
                //获取列表
                List<XSSFShape> shapeList = drawing.getShapes();
                //循环2
                for (XSSFShape shape : shapeList) {
                    //强转
                    XSSFPicture pic = (XSSFPicture) shape;
                    //获取其可能的表格
                    CTMarker ctMarker = pic.getPreferredSize().getFrom();
                    //如果存在改key
                    if (keyMap.containsKey(ctMarker.getCol())) {
                        //获取key
                        String key = keyMap.get(ctMarker.getCol());
                        //尝试获取结果map
                        Map<Integer, PictureData> map = result.getOrDefault(key, new HashMap<>());
                        //组装
                        map.put(ctMarker.getRow() - 1, pic.getPictureData());
                        result.put(key, map);
                    }
                }
            }
        }
        //返回
        return result;
    }




    /**
     * 解析XLSX图片
     *
     * @return
     */
    private static Map<String, Map<Integer, PictureData>> parseSheetPicForXLSX
    (XSSFSheet xssfSheet, Map<Integer, String> keyMap) {
        //初始化结果
        Map<String, Map<Integer, PictureData>> result = new HashMap<>();
        //获取列表
        List<POIXMLDocumentPart> relationList = xssfSheet.getRelations();
        //循环1
        for (POIXMLDocumentPart dr : relationList) {
            //如果是类型
            if (dr instanceof XSSFDrawing) {
                //强转
                XSSFDrawing drawing = (XSSFDrawing) dr;
                //获取列表
                List<XSSFShape> shapeList = drawing.getShapes();
                //循环2
                for (XSSFShape shape : shapeList) {
                    //强转
                    XSSFPicture pic = (XSSFPicture) shape;
                    //获取其可能的表格
                    CTMarker ctMarker = pic.getPreferredSize().getFrom();
                    //如果存在改key
                    if (keyMap.containsKey(ctMarker.getCol())) {
                        //获取key
                        String key = keyMap.get(ctMarker.getCol());
                        //尝试获取结果map
                        Map<Integer, PictureData> map = result.getOrDefault(key, new HashMap<>());
                        //组装
                        map.put(ctMarker.getRow() - 1, pic.getPictureData());
                        result.put(key, map);
                    }
                }
            }
        }
        //返回
        return result;
    }


}
