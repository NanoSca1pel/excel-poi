package com.lht.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;

/**
 * @author lhtao
 * @date 2020/6/4 15:42
 */
public class ExcelWriteTest {

    private static final String PATH = "C:\\Users\\Administrator\\Desktop\\";

    @Test
    public void testWrite03() throws Exception {

        //1.创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("lht测试excel");
        //3.创建一个行
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //第二行
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");

        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));


        FileOutputStream fos = new FileOutputStream(PATH + "lht测试excel03.xls");
        workbook.write(fos);
    }

    @Test
    public void testWrite07() throws Exception {
        //1.创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("lht测试excel");
        //3.创建一个行
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //第二行
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");

        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));


        FileOutputStream fos = new FileOutputStream(PATH + "lht测试excel07.xlsx");
        workbook.write(fos);

    }

    @Test
    public void testWrite03BigData() throws Exception {
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();

        //创建一个表
        Sheet sheet = workbook.createSheet();

        //写入数据
        for(int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10 ; cellNum ++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        System.out.println("---------over---------");
        FileOutputStream fos = new FileOutputStream(PATH + "1.xls");
        workbook.write(fos);
        fos.close();

        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }

    /** 耗时长 */
    @Test
    public void testWrite07BigData() throws Exception {
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook();

        //创建一个表
        Sheet sheet = workbook.createSheet();

        //写入数据
        for(int rowNum = 0; rowNum < 100000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10 ; cellNum ++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        System.out.println("---------over---------");
        FileOutputStream fos = new FileOutputStream(PATH + "1.xlsx");
        workbook.write(fos);
        fos.close();

        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }

    @Test
    public void testWrite07BigDataSuper() throws Exception {
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new SXSSFWorkbook();

        //创建一个表
        Sheet sheet = workbook.createSheet();

        //写入数据
        for(int rowNum = 0; rowNum < 100000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10 ; cellNum ++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        System.out.println("---------over---------");
        FileOutputStream fos = new FileOutputStream(PATH + "1.xlsx");
        workbook.write(fos);
        fos.close();

        //清除临时文件
        ((SXSSFWorkbook) workbook).dispose();

        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }
}
