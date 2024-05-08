package com.ztj.excel;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

/**
 * @author 29027
 */
public class ExcelToWord {
    public static void excelToWord(HttpServletResponse response, XSSFWorkbook workBook, Integer index) {

        XWPFDocument doc = new XWPFDocument();
        CTBody body = doc.getDocument().getBody();
        if (!body.isSetSectPr()) {
            body.addNewSectPr();
        }
        XSSFSheet sheetAt = workBook.getSheetAt(index);
        int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
        int physicalNumberOfCells = sheetAt.getRow(0).getPhysicalNumberOfCells();
        //创建一个表
        XWPFTable table = doc.createTable(physicalNumberOfRows, physicalNumberOfCells);
        CTTblWidth ctTblWidth = table.getCTTbl().addNewTblPr().addNewTblW();
        ctTblWidth.setType(STTblWidth.PCT);
        CTTblGrid ctTblGrid = table.getCTTbl().addNewTblGrid();
        ctTblWidth.setW(new BigInteger("4900"));

        for (int e=0;e<physicalNumberOfCells;e++){
            ctTblGrid.addNewGridCol().setW(new BigInteger(String.valueOf(12000/physicalNumberOfCells)));
        }
        for (int r = 0; r < sheetAt.getPhysicalNumberOfRows(); r++) {

            for (int c = 0; c < sheetAt.getRow(r).getPhysicalNumberOfCells(); c++) {
                List<XWPFParagraph> paragraphs = table.getRow(r).getCell(c).getParagraphs();
                for (Integer m = 0; m < paragraphs.size(); m++) {

                    XWPFParagraph xwpfParagraph = paragraphs.get(m);
                    //设置局中
                    xwpfParagraph.setAlignment(ParagraphAlignment.CENTER);
                    xwpfParagraph.setVerticalAlignment(TextAlignment.CENTER);
                    System.out.println("====================================");
                    System.out.println(xwpfParagraph.getText());
                    System.out.println(xwpfParagraph);
                }

                widthCellsAcrossRow(table, r, c, 1200);
                if (isMergedRegion(sheetAt, r, c)) {
                    int[] span = getMergedSpan(sheetAt, r, c);
                    if (span[0] == 1 && span[1] == 1) {//忽略合并过的单元格
                        continue;
                    }
                    mergeCellsHorizontal(table, r, span[0] - 1, span[1] - 1, ExcelUtil.getCellStringVal(sheetAt.getRow(r).getCell(c)));
                    c = c + span[1] - 1;//合并过的列直接跳过
                }
                table.getRow(r).getCell(c).setText(ExcelUtil.getCellStringVal(sheetAt.getRow(r).getCell(c)));


            }
        }


        ServletOutputStream out;
        try {
            response.setHeader("Content-Disposition", "inline;fileName=" + "测试" + ".doc");
            response.setContentType("application/pdf");
            out = response.getOutputStream();
            doc.write(out);
            out.flush();
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * @Description: 跨列合并
     * table要合并单元格的表格
     * row要合并哪一行的单元格
     * fromCell开始合并的单元格
     * toCell合并到哪一个单元格
     * text 合并文本
     */
    public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell, String text) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
            cell.setText(text);
        }
    }

    /**
     * * 判断单元格是否是合并单元格
     * * @param sheet
     * * @param row
     * * @param column
     * * @return
     */
    private static boolean isMergedRegion(XSSFSheet sheet, int row, int column) {
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


    private static int[] getMergedSpan(XSSFSheet sheet, int row, int column) {
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

    private static void widthCellsAcrossRow(XWPFTable table, int rowNum, int colNum, int width) {
        XWPFTableCell cell = table.getRow(rowNum).getCell(colNum);
        if (cell.getCTTc().getTcPr() == null) {
            cell.getCTTc().addNewTcPr();
        }
        if (cell.getCTTc().getTcPr().getTcW() == null) {
            cell.getCTTc().getTcPr().addNewTcW();
        }
        cell.getCTTc().getTcPr().getTcW().setW(BigInteger.valueOf((long) width));
    }
}
