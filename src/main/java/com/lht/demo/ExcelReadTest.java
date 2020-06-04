package com.lht.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;

/**
 * @author lhtao
 * @date 2020/6/4 15:42
 */
public class ExcelReadTest {

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
}
