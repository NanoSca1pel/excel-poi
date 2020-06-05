package com.lht.demo;

import com.sun.org.apache.bcel.internal.generic.NEW;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.Date;

/**
 * @author lhtao
 * @date 2020/6/4 15:42
 */
public class PoiExcelReadTest {

    private static final String PATH = "C:\\Users\\Administrator\\Desktop\\";

    @Test
    public void testRead03() throws Exception{

        //获取文件流
        FileInputStream fis = new FileInputStream(PATH + "1.xls");

        //1.创建一个工作簿。使用excel能操作的这里都可以操作
        Workbook workbook = new HSSFWorkbook(fis);
        //2.得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3.得到行
        Row row = sheet.getRow(0);
        //4.得到列
        Cell cell = row.getCell(0);

        //读取值的时候，一定要注意类型！
        System.out.println(cell.getNumericCellValue());
        fis.close();
    }

    @Test
    public void testRead07() throws Exception{

        //获取文件流
        FileInputStream fis = new FileInputStream(PATH + "1.xlsx");

        //1.创建一个工作簿。使用excel能操作的这里都可以操作
        Workbook workbook = new XSSFWorkbook(fis);
        //2.得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3.得到行
        Row row = sheet.getRow(0);
        //4.得到列
        Cell cell = row.getCell(0);

        //读取值的时候，一定要注意类型！
        System.out.println(cell.getNumericCellValue());
        fis.close();
    }

    @Test
    public void testCellType03() throws Exception {

        //1.获取文件流
        FileInputStream fis = new FileInputStream(PATH + "1.xls");
        //2，创建一个工作簿
        Workbook workbook = new HSSFWorkbook(fis);
        //3.获取第一个标签页
        Sheet sheet = workbook.getSheetAt(0);
        //4.获取第一行的标题内容
        Row rowTitle = sheet.getRow(0);

        /*if (rowTitle != null) {
            //获取一行中的单元格个数
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    //获取单元格值的元素类型
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.println(cellValue + " | ");
                }
            }
            System.out.println();
        }*/

        //获取表中的内容
        //首先读取每行
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 0; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {
                //接着开始读取行中每列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {

                    Cell cell = rowData.getCell(cellNum);
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        String cellValue = "";

                        switch(cellType) {
                            case Cell.CELL_TYPE_NUMERIC:  //数字(日期、普通数字)使用改这个类型
                                System.out.println("【NUMERIC】");
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    System.out.println("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd HH:mm:ss");
                                } else {
                                    //如果不是日期格式，防止数字过长
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case Cell.CELL_TYPE_STRING:
                                System.out.println("【STRING】");
                                cellValue = cell.getStringCellValue();
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                System.out.println("【BLANK】");
                                break;
                            case Cell.CELL_TYPE_BOOLEAN:
                                System.out.println("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                System.out.println("【ERROR】");
                                cellValue = String.valueOf(cell.getErrorCellValue());
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        fis.close();
    }

    @Test
    public void testCellType07() throws Exception {

        //1.获取文件流
        FileInputStream fis = new FileInputStream(PATH + "1.xlsx");
        //2，创建一个工作簿
        Workbook workbook = new XSSFWorkbook(fis);
        //3.获取第一个标签页
        Sheet sheet = workbook.getSheetAt(0);
        //4.获取第一行的标题内容
        Row rowTitle = sheet.getRow(0);

        /*if (rowTitle != null) {
            //获取一行中的单元格个数
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    //获取单元格值的元素类型
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.println(cellValue + " | ");
                }
            }
            System.out.println();
        }*/

        //获取表中的内容
        //首先读取每行
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 0; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {
                //接着开始读取行中每列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {

                    Cell cell = rowData.getCell(cellNum);
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        String cellValue = "";

                        switch(cellType) {
                            case Cell.CELL_TYPE_NUMERIC:  //数字(日期、普通数字)使用改这个类型
                                System.out.println("【NUMERIC】");
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    System.out.println("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd HH:mm:ss");
                                } else {
                                    //如果不是日期格式，防止数字过长
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case Cell.CELL_TYPE_STRING:
                                System.out.println("【STRING】");
                                cellValue = cell.getStringCellValue();
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                System.out.println("【BLANK】");
                                break;
                            case Cell.CELL_TYPE_BOOLEAN:
                                System.out.println("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                System.out.println("【ERROR】");
                                cellValue = String.valueOf(cell.getErrorCellValue());
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        fis.close();
    }

    @Test
    public void testFormula03() throws Exception {
        FileInputStream fis = new FileInputStream(PATH + "1.xls");
        Workbook workbook = new HSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(1);
        Cell cell = row.getCell(0);

        //获取计算公式
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook)workbook);

        //输出单元格的内容
        int cellType = cell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_FORMULA: //公式
                String formula = cell.getCellFormula();
                System.out.println(formula);

                //计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }
    }
}
