package com.ztj.excel;


import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;


/**
 * @author 29027
 */
public class ExcelToPDF {

    public static void excelToPDF(HttpServletResponse response, XSSFWorkbook workBook, String pdfName) throws Exception {

        //       也可以改成直接读取Excel文件目录就可以直接Excel转PDF
//        InputStream inputStream = new FileInputStream(Excel文件绝对路径);
//        HSSFWorkbook workBook = new HSSFWorkbook(inputStream);
        XSSFSheet sheet = workBook.getSheetAt(0);
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        System.out.println("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa");
        Rectangle pageSize = new Rectangle(842, 595);
        pageSize.rotate();
        Document document = new Document(pageSize);//设置pdf纸张大小

        PdfWriter writer = PdfWriter.getInstance(document, stream);
        document.setMargins(0, 0, 15, 15);//设置页边距
        document.open();
        float[] widths = getColWidth(sheet);//获取excel每列宽度占比

        PdfPTable table = new PdfPTable(widths);//初始化pdf中每列的宽度
        table.setWidthPercentage(88);

        int colCount = widths.length;

        BaseFont baseFont = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",
                BaseFont.NOT_EMBEDDED);//设置基本字体
         //开始遍历excel内容并绘制pdf
        for (int r = sheet.getFirstRowNum(); r < sheet.getPhysicalNumberOfRows(); r++) {
            XSSFRow row = sheet.getRow(r);
            if (row != null) {
                if (!ExcelUtil.isRowEmpty(row)){
                for (int c = row.getFirstCellNum(); (c < row.getLastCellNum() || c < colCount) && c > -1; c++) {
                    if (c >= row.getPhysicalNumberOfCells()) {
                        PdfPCell pCell = new PdfPCell(new Phrase(""));
                        pCell.setBorder(0);
                        table.addCell(pCell);
                        continue;
                    }
                    XSSFCell excelCell = row.getCell(c);
                    String value = "";

                    if (excelCell != null) {
                        value = excelCell.toString().trim();
                        if (value != null && value.length() != 0) {
                            String dataFormat = excelCell.getCellStyle().getDataFormatString();//获取excel单元格数据显示样式
                            if (dataFormat != "General" && dataFormat != "@") {
                                try {
                                    String numStyle = getNumStyle(dataFormat);
                                    value = numFormat(numStyle, excelCell.getNumericCellValue());
                                } catch (Exception e) {

                                }
                            }
                        }
                    }

//                    XSSFFont font = excelCell.getCellStyle().getFont();

                    // Font.BOLD : 正体大 NORMAL 正体小   ITALIC :斜体小  BOLDITALIC :合并BOLD和ITALIC
                    Font pdFont = new Font(baseFont, 12,
                           Font.NORMAL, BaseColor.BLACK);//设置单元格字体

                    PdfPCell pCell = new PdfPCell(new Phrase(value, pdFont));
//                    pCell.setBorder(0);

//                    boolean hasBorder = hasBorder(excelCell);
//                    if (!hasBorder) {
//                        pCell.setBorder(4);
//                    }
//                    pCell.setHorizontalAlignment();
//                    pCell.setVerticalAlignment(getVerAglin(excelCell.getCellStyle().getVerticalAlignment()));
                    pCell.setHorizontalAlignment(HorizontalAlignment.LEFT.getCode());
                    pCell.setVerticalAlignment(VerticalAlignment.CENTER.getCode());
                    pCell.setMinimumHeight(row.getHeightInPoints());
                    if (isMergedRegion(sheet, r, c)) {
                        int[] span = getMergedSpan(sheet, r, c);
                        if (span[0] == 1 && span[1] == 1) {//忽略合并过的单元格
                            continue;
                        }
                        pCell.setRowspan(span[0]);
                        pCell.setColspan(span[1]);
                        c = c + span[1] - 1;//合并过的列直接跳过
                    }

                    table.addCell(pCell);

                }
            } else {
                PdfPCell pCell = new PdfPCell(new Phrase(""));
                pCell.setBorder(0);
                pCell.setMinimumHeight(13);
                table.addCell(pCell);
            }}
        }
        document.add(table);
        document.close();

        byte[] pdfByte = stream.toByteArray();
        stream.close();

        ServletOutputStream out;
        try {
            response.setHeader("Content-Disposition", "inline;fileName=" + pdfName + ".pdf");
            response.setContentType("application/pdf");
            out = response.getOutputStream();
            out.write(pdfByte);
            out.flush();
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }


    /**
     *      * 获取excel中每列宽度的占比
     *      * @param sheet
     *      * @return
     *
     */
    private static float[] getColWidth( XSSFSheet sheet) {
        int rowNum = getMaxColRowNum(sheet);
        XSSFRow row = sheet.getRow(rowNum);
        int cellCount = row.getPhysicalNumberOfCells();
        int[] colWidths = new int[cellCount];
        int sum = 0;

        for (int i = row.getFirstCellNum(); i < cellCount; i++) {
            XSSFCell cell = row.getCell(i);
            if (cell != null) {
                colWidths[i] = sheet.getColumnWidth(i);
                sum += sheet.getColumnWidth(i);
            }
        }

        float[] colWidthPer = new float[cellCount];
        for (int i = row.getFirstCellNum(); i < cellCount; i++) {
            colWidthPer[i] = (float) colWidths[i] / sum * 100;
        }
        return colWidthPer;
    }


    /**
     *      * 获取excel中列数最多的行号
     *      * @param sheet
     *      * @return
     *
     */
    private static int getMaxColRowNum( XSSFSheet sheet) {
        int rowNum = 0;
        int maxCol = 0;
        for (int r = sheet.getFirstRowNum(); r < sheet.getPhysicalNumberOfRows(); r++) {
            XSSFRow row = sheet.getRow(r);
            if (row != null && maxCol < row.getPhysicalNumberOfCells()) {
                maxCol = row.getPhysicalNumberOfCells();
                rowNum = r;
            }
        }
        return rowNum;
    }


    /**
     *      * 获取excel单元格数据显示格式
     *      * @param dataFormat
     *      * @return
     *      * @throws Exception
     *
     */
    private static String getNumStyle(String dataFormat) throws Exception {
        if (dataFormat == null || dataFormat.length() == 0) {
            throw new Exception("");
        }
        if (dataFormat.indexOf("%") > -1) {
            return dataFormat;
        } else {
            return dataFormat.substring(0, dataFormat.length() - 2);
        }

    }


    /**
     *      * 判断单元格是否是合并单元格
     *      * @param sheet
     *      * @param row
     *      * @param column
     *      * @return
     *
     */
    private static boolean isMergedRegion( XSSFSheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }


    /**
     *      * 计算合并单元格合并的跨行跨列数
     *      * @param sheet
     *      * @param row
     *      * @param column
     *      * @return
     *
     */
    private static int[] getMergedSpan( XSSFSheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        int[] span = {1, 1};
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (firstColumn == column && firstRow == row) {
                span[0] = lastRow - firstRow + 1;
                span[1] = lastColumn - firstColumn + 1;
                break;
            }
        }
        return span;
    }


    /**
     *      * 判断excel单元格是否有边框
     *      * @param excelCell
     *      * @return
     *
     */
    private static boolean hasBorder(   XSSFCell excelCell ) {
        short top = excelCell.getCellStyle().getBorderTopEnum().getCode();
        short bottom = excelCell.getCellStyle().getBorderBottomEnum().getCode();
        short left = excelCell.getCellStyle().getBorderLeftEnum().getCode();
        short right = excelCell.getCellStyle().getBorderRightEnum().getCode();
        return top + bottom + left + right > 2;
    }

    /**
     *      * excel水平对齐方式映射到pdf水平对齐方式
     *      * @param aglin
     *      * @return
     *
     */
    private static int getHorAglin(int aglin) {
        switch (aglin) {
            case 2:
                return Element.ALIGN_CENTER;
            case 3:
                return Element.ALIGN_RIGHT;
            case 1:
                return Element.ALIGN_LEFT;
            default:
                return Element.ALIGN_CENTER;
        }
    }


    /**
     *      * excel垂直对齐方式映射到pdf对齐方式
     *      * @param aglin
     *      * @return
     *
     */
    private static int getVerAglin(int aglin) {
        switch (aglin) {
            case 1:
                return Element.ALIGN_MIDDLE;
            case 2:
                return Element.ALIGN_BOTTOM;
            case 3:
                return Element.ALIGN_TOP;
            default:
                return Element.ALIGN_MIDDLE;
        }
    }

    /**
     *      * 格式化数字
     *      * @param pattern
     *      * @param num
     *      * @return
     *
     */
    private static String numFormat(String pattern, double num) {
        DecimalFormat format = new DecimalFormat(pattern);
        return format.format(num);
    }



}
