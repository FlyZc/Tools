package com.zhouchi.tool;

/**
 * @Project: Tools
 * @Description: 筛选TD34数据
 * @Author: ChiZhou
 * @Date: 2020-08-12 08:44
 */

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashSet;
import java.util.Set;

public class TD34 {
    public static void main(String[] args) throws Exception {
        File fileInfo = new File("E:\\Daily\\Information.txt");
        FileOutputStream outputStreamInfo = new FileOutputStream(fileInfo);
        OutputStreamWriter streamWriterInfo = new OutputStreamWriter(outputStreamInfo);

        //读取需统计用户的数据
        InputStream streamUser = new FileInputStream("E:\\Daily\\TD34.xlsx");
        if (streamUser == null) {
            streamWriterInfo.append("TD34全网通话记录<TD34.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUser = new XSSFWorkbook(streamUser);
        //获取excel表的第一个sheet
        XSSFSheet sheetUser = wbUser.getSheetAt(0);
        if (sheetUser == null) {
            return;
        }

        //读取TD34 DH方向用户的数据
        InputStream streamUserDH = new FileInputStream("E:\\Daily\\TDDH34.xlsx");
        if (streamUserDH == null) {
            streamWriterInfo.append("TD34 DH方向用户名单<TDDH34.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserDH = new XSSFWorkbook(streamUserDH);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserDH = wbUserDH.getSheetAt(0);
        if (sheetUserDH == null) {
            return;
        }

        //读取TD34 NH方向用户的数据
        InputStream streamUserNH = new FileInputStream("E:\\Daily\\TDNH34.xlsx");
        if (streamUserNH == null) {
            streamWriterInfo.append("TD34 NH方向用户名单<TDNH34.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserNH = new XSSFWorkbook(streamUserNH);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserNH = wbUserNH.getSheetAt(0);
        if (sheetUserNH == null) {
            return;
        }

        //读取FD NH方向用户的数据
        InputStream streamUserTH = new FileInputStream("E:\\Daily\\TDTH34.xlsx");
        if (streamUserTH == null) {
            streamWriterInfo.append("TD34 TH方向用户名单<TDTH34.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserTH = new XSSFWorkbook(streamUserTH);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserTH = wbUserTH.getSheetAt(0);
        if (sheetUserTH == null) {
            return;
        }

        int countUser = 0;
        Set<String> userAll = new HashSet<String>(sheetUser.getLastRowNum());


        Set<String> userDH = new HashSet<String>(sheetUserDH.getLastRowNum());
        Set<String> userDHResult = new HashSet<String>(sheetUserDH.getLastRowNum());

        Set<String> userNH = new HashSet<String>(sheetUserNH.getLastRowNum());
        Set<String> userNHResult = new HashSet<String>(sheetUserNH.getLastRowNum());

        Set<String> userTH = new HashSet<String>(sheetUserTH.getLastRowNum());
        Set<String> userTHResult = new HashSet<String>(sheetUserTH.getLastRowNum());

        File fileAll = new File("E:\\Daily\\TD34.txt");
        File fileDH = new File("E:\\Daily\\TDDH34.txt");
        File fileTH = new File("E:\\Daily\\TDTH34.txt");
        File fileNH = new File("E:\\Daily\\TDNH34.txt");

        //获取ZB用户站列表
        readUserSet(sheetUserNH, userNH);
        readUserSet(sheetUserTH, userTH);
        readUserSet(sheetUserDH, userDH);

        //获取所有通信用户站列表;
        String path = "E:\\Daily\\TD34.xlsx";

        calAllRate(path, userAll,fileAll,"TD34全网呼通率统计");
        calIRate(path, userAll, userDH, userDHResult, fileDH,"TD34 DH方向呼通率统计");
        calIRate(path, userAll, userNH, userNHResult, fileNH,"TD34 NH方向呼通率统计");
        calIRate(path, userAll, userTH, userTHResult, fileTH,"TD34 TH方向呼通率统计");
    }

    public static String readCellMethod(XSSFCell cell) {
        return cell.getStringCellValue();
    }

    public static void calIRate(String allPath, Set<String> userAll, Set<String> user, Set<String> userResult, File file, String description) {
        try {
            //读取所有通话记录的信息
            InputStream stream = new FileInputStream(allPath);
            if (stream == null) {

            } else {
                XSSFWorkbook wb = new XSSFWorkbook(stream);

                //获取excel表的第一个sheet
                XSSFSheet sheet = wb.getSheetAt(0);

                if (sheet == null) {
                    return;
                }

                //用来记录筛选出的用户站通话成功的数量
                int successCount = 0;

                //遍历该sheet的行
                for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    XSSFRow row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }

                    //筛选主叫地址
                    int cloumn1 = 1;
                    XSSFCell callAddress = row.getCell(cloumn1);

                    //呼叫结果
                    int column2 = 6;
                    XSSFCell result = row.getCell(column2);

                    if (callAddress == null) {
                        continue;
                    } else {
                        if (user.contains(callAddress.toString())) {
                            if (("成功".equals(result.toString()))) {
                                userResult.add(callAddress.toString());
                                System.out.println("34 " + callAddress.toString());
                                successCount++;
                            }
                        }
                    }

                    try (FileOutputStream outputStream = new FileOutputStream(file);
                         OutputStreamWriter streamWriter = new OutputStreamWriter(outputStream)) {
                        streamWriter.append("========================" + description + "========================\n");
                        streamWriter.append("\n********重保********\n");
                        streamWriter.append("总呼叫数：" + successCount + "\n");
                        streamWriter.append("呼叫成功数：" + successCount + "\n");
                        if (successCount != 0) {
                            streamWriter.append("呼通率：100%\n");
                        } else {
                            streamWriter.append("呼通率：\n");
                        }

                        streamWriter.append("通信用户站数量：" + userResult.size() + "\n");
                        streamWriter.append("通信用户站列表：");
                        int count = 0;
                        for (String str : userResult) {
                            count++;
                            if (count % 10 == 0) {
                                streamWriter.append("\n");
                            }
                            streamWriter.append(str + "  ");
                        }
                        streamWriter.append("\n========================" + description + "========================\n");
                        stream.close();
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void calAllRate(String allPath, Set<String> userAll, File file, String description) {
        try {
            //读取所有通话记录的信息
            InputStream stream = new FileInputStream(allPath);
            if (stream == null) {

            } else {
                XSSFWorkbook wb = new XSSFWorkbook(stream);

                //获取excel表的第一个sheet
                XSSFSheet sheet = wb.getSheetAt(0);
                if (sheet == null) {
                    return;
                }
                
                //所有通话记录数量
                int allRecordLines = sheet.getLastRowNum() - 1;

                //用来记录筛选出的用户站通话成功的数量
                int successCountAll = 0;

                //遍历该sheet的行
                for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    XSSFRow row = sheet.getRow(rowNum);
                    //筛选主叫地址
                    int cloumn1 = 1;
                    XSSFCell callAddress = row.getCell(cloumn1);

                    //呼叫结果
                    int column2 = 6;
                    XSSFCell result = row.getCell(column2);

                    if (row == null) {
                        continue;
                    }
                    if (rowNum > 1 && ("成功".equals(result.toString()))) {
                        userAll.add(row.getCell(1).toString());
                    }

                    if (callAddress == null) {
                        continue;
                    } else {
                        if (("成功".equals(result.toString()))) {
                            successCountAll++;
                        }
                    }
                    
                    try (FileOutputStream outputStream = new FileOutputStream(file);
                         OutputStreamWriter streamWriter = new OutputStreamWriter(outputStream)) {

                        streamWriter.append("========================" + description + "========================\n");
                        streamWriter.append(">>>>>>>>全网<<<<<<<<\n");
                        streamWriter.append("总呼叫数：" + successCountAll + "\n");
                        streamWriter.append("呼叫成功数：" + successCountAll + "\n");
                        if (successCountAll != 0) {
                            streamWriter.append("呼通率：100%\n");
                        } else {
                            streamWriter.append("呼通率：\n");
                        }

                        streamWriter.append("通信用户站数量：" + userAll.size() + "\n");
                        streamWriter.append("通信用户站列表：\n");
                        int count = 0;
                        for (String str : userAll) {
                            count++;
                            if (count % 10 == 0) {
                                streamWriter.append("\n");
                            }
                            streamWriter.append(str + "  ");
                        }
                        streamWriter.append("\n========================" + description + "========================\n");
                        stream.close();
                    } catch (FileNotFoundException ex) {
                        ex.printStackTrace();
                    } catch (IOException ex) {
                        ex.printStackTrace();
                    }
                }
            }
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    public static void readUserSet(XSSFSheet sheetUser, Set<String> user) {
        //遍历该sheet的行，读取用户数据，存入数组中。
        for (int rowNum = 0; rowNum <= sheetUser.getLastRowNum(); rowNum++) {
            XSSFRow row = sheetUser.getRow(rowNum);
            if (row == null) {
                continue;
            } else {
                row.getCell(0).setCellType(CellType.STRING);
                user.add(String.valueOf(row.getCell(0).getStringCellValue()));
            }
        }
    }
}