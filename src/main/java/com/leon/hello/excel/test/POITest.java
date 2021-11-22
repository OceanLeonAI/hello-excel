package com.leon.hello.excel.test;

import com.leon.hello.excel.util.FileUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @PROJECT_NAME: hello-excel
 * @CLASS_NAME: SimpleTest
 * @AUTHOR: OceanLeonAI
 * @CREATED_DATE: 2021/11/22 10:41
 * @Version 1.0
 * @DESCRIPTION:
 **/
public class POITest {

    // 写入文件目录为当前项目编译后目录
    static String FILE_PATH = FileUtil.getPath();

    public static void main(String[] args) throws IOException {
//        write_03();
        read_03();
//        write_07();
//        read_07();
    }

    //03版本 ，向excel中写入数据
    public static void write_03() throws IOException {
        //1.创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //2.创建表名
        Sheet sheet = workbook.createSheet("03版_第一张表");
        //3.创建行
        Row row0 = sheet.createRow(0);
        //4.创建单元格
        Cell cell = row0.createCell(0);
        //5.写入数据
        cell.setCellValue("这是第一行一列的格子");
        //6.创建流用于输出
        FileOutputStream fileOutputStream = new FileOutputStream(FILE_PATH + "03版本表.xls");
        //7.输出
        workbook.write(fileOutputStream);

        System.out.println("03版本表已经生成");

    }

    //07版本 ，向excel中写入数据
    public static void write_07() throws IOException {
        //1.创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //2.创建表名
        Sheet sheet = workbook.createSheet("07版_第一张表");
        //3.创建行
        Row row0 = sheet.createRow(0);
        //4.创建单元格
        Cell cell = row0.createCell(0);
        //5.写入数据
        cell.setCellValue("这是第一行一列的格子");
        //6.创建流用于输出
        FileOutputStream fileOutputStream = new FileOutputStream(FILE_PATH + "07版本表.xlsx");
        //7.输出
        workbook.write(fileOutputStream);

        System.out.println("07版本表已经生成");
    }

    //03版本 ，读excel中的数据
    public static void read_03() throws IOException {
        //1.创建流用于读取
        FileInputStream fileInputStream = new FileInputStream(FILE_PATH + "03版本表.xls");
        //2.工作簿用于接收
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        //3.获取表
        Sheet sheet = workbook.getSheetAt(0);
        //4.获取行
        Row row = sheet.getRow(0);
        //5.获取单元格
        Cell cell = row.getCell(0);
        //6.获取单元格中的数据,并输出
        System.out.println(cell.getStringCellValue());

        System.out.println("获取到03版本中的数据");
    }

    //07版本 ，读excel中的数据
    public static void read_07() throws IOException {
        //1.创建流用于读取
        FileInputStream fileInputStream = new FileInputStream(FILE_PATH + "07版本表.xlsx");
        //2.工作簿用于接收
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        //3.获取表
        Sheet sheet = workbook.getSheetAt(0);
        //4.获取行
        Row row = sheet.getRow(0);
        //5.获取单元格
        Cell cell = row.getCell(0);
        //6.获取单元格中的数据,并输出
        System.out.println(cell.getStringCellValue());

        System.out.println("获取到07版本中的数据");
    }

}
