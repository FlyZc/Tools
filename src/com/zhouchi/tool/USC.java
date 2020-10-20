package com.zhouchi.tool;

/**
 * @Project: Tools
 * @Description: 筛选USC数据
 * @Author: ChiZhou
 * @Date: 2020-08-12 08:44
 */

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.HashSet;
import java.util.Set;


public class USC {
    public static void main(String[] args) throws Exception {
        File fileInfo = new File("G:\\Information.txt");
        FileOutputStream outputStreamInfo = new FileOutputStream(fileInfo);
        OutputStreamWriter streamWriterInfo = new OutputStreamWriter(outputStreamInfo);
        //读取需统计用户的数据
        InputStream streamUser = new FileInputStream("G:\\UserUSC.xlsx");
        if (streamUser == null) {
            streamWriterInfo.append("USC重保用户名单<UserUSC.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUser = new XSSFWorkbook(streamUser);

        //获取excel表的第一个sheet
        XSSFSheet sheetUser = wbUser.getSheetAt(0);
        String sheetNameUser = wbUser.getSheetName(0);

        if (sheetUser == null) {
            return;
        }
        int countUser = 0;
        Set<String> user = new HashSet<String>(sheetUser.getLastRowNum());
        Set<String> userAll = new HashSet<String>(sheetUser.getLastRowNum());
        Set<String> userResult = new HashSet<String>(sheetUser.getLastRowNum());
        //遍历该sheet的行，读取用户数据，存入数组中。
        for (int rowNum = 0; rowNum <= sheetUser.getLastRowNum(); rowNum++) {
            XSSFRow row = sheetUser.getRow(rowNum);
            if (row == null) {
                continue;
            } else {
                user.add(row.getCell(0).toString());
            }
        }


        //读取通信记录表的内容
        InputStream stream = new FileInputStream("G:\\USC.xlsx");
        if (streamUser == null) {
            streamWriterInfo.append("USC通话记录文件<USC.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wb = new XSSFWorkbook(stream);

        //获取excel表的第一个sheet
        XSSFSheet sheet = wb.getSheetAt(0);
        String sheetName = wb.getSheetName(0);

        if (sheet == null) {
            return;
        }

        //所有通话记录数量
        int allRecordLines = sheet.getLastRowNum() - 1;
        //用来记录筛选出的用户站所有拨打电话的数量
        int allUserCount = 0;

        //遍历该sheet的行
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            XSSFRow row = sheet.getRow(rowNum);
            if (row == null) {
                continue;
            }
            if (rowNum > 1) {
                userAll.add(row.getCell(1).toString());
            }


            //筛选主叫地址
            int cloumn1 = 1;
            XSSFCell callAddress = row.getCell(cloumn1);



            if (callAddress == null ) {
                continue;
            } else {
                if (user.contains(callAddress.toString())) {
                    userResult.add(callAddress.toString());
                    allUserCount++;
                }
            }
        }

        File file = new File("G:\\USC.txt");
        try(FileOutputStream outputStream = new FileOutputStream(file);
            OutputStreamWriter streamWriter = new OutputStreamWriter(outputStream)) {
            streamWriter.write("========================USC========================\n");
            streamWriter.append(">>>>>>>>全网<<<<<<<<\n");
            streamWriter.append("通信用户站数量：" + userAll.size() + "\n");
            streamWriter.append("通信用户站列表：\n");
            for(String str : userAll){
                streamWriter.append(str + "  ");
            }

            streamWriter.append("\n********重保********\n");
            streamWriter.append("总呼叫数：" + allUserCount + "\n");
            streamWriter.append("通信用户站数量：" + userResult.size() + "\n");
            streamWriter.append("通信用户站列表：");
            for(String str : userResult){
                streamWriter.append(str + "  ");
            }
            streamWriter.append("\n========================USC========================\n");
            stream.close();
        }



        stream.close();
    }

    public static String readCellMethod(XSSFCell cell) {
        return cell.getStringCellValue();
    }
}