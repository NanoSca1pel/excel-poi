package com.lht.demo;

import com.alibaba.excel.EasyExcel;
import com.lht.entity.Student;
import com.lht.entity.StudentListener;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormat;
import org.junit.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * @author lhtao
 * @date 2020/6/5 10:09
 */
public class EasyExcelWriteTest {

    private static final String PATH = "C:\\Users\\Administrator\\Desktop\\";

    private static List<Student> data() {
        List<Student> list = new ArrayList<>(10);
        for (int i = 1; i <= 10; i++) {
            Student student = new Student();
            student.setName("王文源" + i);
            student.setBirth(DateTime.parse("1995-01-03 13:00:41", DateTimeFormat.forPattern("yyyy-MM-dd HH:mm:ss")).toDate());
            student.setAge(25 + i);
            student.setGender("男");
            list.add(student);
        }
        return list;
    }

    /**
     * 最简单的写
     * <p>1. 创建excel对应的实体对象 参照{@link com.lht.entity.Student}
     * <p>2. 直接写即可
     */
    @Test
    public void simpleWrite() {
        String fileName = PATH + "2.xlsx";
        // 这里 需要指定写用哪个class去读，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
        EasyExcel.write(fileName, Student.class).sheet("模板").doWrite(data());
    }


    /**
     * 最简单的读
     * <p>1. 创建excel对应的实体对象 参照{@link Student}
     * <p>2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照{@link StudentListener}
     * <p>3. 直接读即可
     */
    @Test
    public void simpleRead() {
        String fileName = PATH + "2.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, Student.class, new StudentListener()).sheet().doRead();
    }
}
