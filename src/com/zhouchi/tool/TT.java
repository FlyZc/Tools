package com.zhouchi.tool;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.HashSet;
import java.util.Set;

/**
 * @Project: Tools
 * @Description: tt呼通率和分组数据统计
 * @Author: ChiZhou
 * @Date: 2021-01-25 14:47
 */

public class TT {
    public static void main(String[] args) throws Exception {
        File fileInfo = new File("E:\\Daily\\Information.txt");
        FileOutputStream outputStreamInfo = new FileOutputStream(fileInfo);
        OutputStreamWriter streamWriterInfo = new OutputStreamWriter(outputStreamInfo);

        //读取需统计的话音数据
        InputStream streamDialog = new FileInputStream("E:\\Daily\\TT1.xlsx");
        //读取需统计的分组数据
        InputStream streamData = new FileInputStream("E:\\Daily\\TT2.xlsx");

        if (streamDialog == null) {
            streamWriterInfo.append("TT系统通话记录<TT1.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }

        if (streamData == null) {
            streamWriterInfo.append("TT系统分组记录<TT2.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }

        XSSFWorkbook wbDialog = new XSSFWorkbook(streamDialog);
        //获取excel表的第一个sheet
        XSSFSheet sheetDialog = wbDialog.getSheetAt(0);
        if (sheetDialog == null) {
            return;
        }

        XSSFWorkbook wbData = new XSSFWorkbook(streamData);
        //获取excel表的第一个sheet
        XSSFSheet sheetData = wbData.getSheetAt(0);
        if (sheetData == null) {
            return;
        }

        Set<String> userAll = new HashSet<String>(sheetDialog.getLastRowNum());
        File fileAll = new File("E:\\Daily\\TT1.txt");
        //获取话音记录文件
        String path = "E:\\Daily\\TT1.xlsx";
        calAllRate(path, userAll,fileAll,"TT话音数据统计");

        Set<String> userData = new HashSet<String>(sheetData.getLastRowNum());
        File fileData = new File("E:\\Daily\\TT2.txt");
        //获取分组数据记录文件
        String pathData = "E:\\Daily\\TT2.xlsx";
        calData(pathData, userData,fileData,"TT分组数据统计");
    }

    public static String readCellMethod(XSSFCell cell) {
        return cell.getStringCellValue();
    }

    public static void calData(String allPath, Set<String> userAll, File file, String description) {
        try {
            //读取所有分组记录的信息
            InputStream stream = new FileInputStream(allPath);
            if (stream == null) {
                //录入提示信息
            } else {
                XSSFWorkbook wb = new XSSFWorkbook(stream);
                //获取excel表的第一个sheet
                XSSFSheet sheet = wb.getSheetAt(0);
                if (sheet == null) {
                    return;
                }
                //用来记录筛选出的用户站所有拨打电话的数量
                int allRecordLines = sheet.getLastRowNum();
                System.out.println(allRecordLines);
                //用来记录筛选出的用户站通话成功的数量
                double sendAll = 0.0;
                double receiveAll = 0.0;

                //遍历该sheet的行
                for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    XSSFRow row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }
                    //筛选终端号
                    int cloumn = 3;
                    XSSFCell callAddress = row.getCell(cloumn);

                    int send = 4;
                    XSSFCell cloumnSend = row.getCell(send);

                    int receive = 5;
                    XSSFCell cloumnReceive = row.getCell(receive);

                    if (rowNum > 0 && (0 != Double.valueOf(cloumnSend.toString()) || 0 != Double.valueOf(cloumnReceive.toString()))) {
                        userAll.add(row.getCell(3).toString());
                        sendAll += Double.valueOf(cloumnSend.toString());
                        receiveAll += Double.valueOf(cloumnReceive.toString());
                    }

                    try (FileOutputStream outputStream = new FileOutputStream(file);
                         OutputStreamWriter streamWriter = new OutputStreamWriter(outputStream)) {

                        streamWriter.append("========================" + description + "========================\n");
                        streamWriter.append(">>>>>>>>系统分组数据<<<<<<<<\n");
                        streamWriter.append("发送数据总量：" + sendAll + "Byte" + "\n");
                        streamWriter.append("发送数据总量：" + sendAll / 1024 + " M" + "\n");
                        streamWriter.append("接收数据总量：" + receiveAll + "Byte" + "\n");
                        streamWriter.append("接收数据总量：" + receiveAll / 1024 + "M" + "\n");
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

    public static void calAllRate(String allPath, Set<String> userAll, File file, String description) {
        try {
            //读取所有通话记录的信息
            InputStream stream = new FileInputStream(allPath);
            if (stream == null) {
                //录入提示信息
            } else {
                XSSFWorkbook wb = new XSSFWorkbook(stream);
                //获取excel表的第一个sheet
                XSSFSheet sheet = wb.getSheetAt(0);
                if (sheet == null) {
                    return;
                }
                //用来记录筛选出的用户站所有拨打电话的数量
                int allRecordLines = sheet.getLastRowNum();
                System.out.println(allRecordLines);
                //用来记录筛选出的用户站通话成功的数量
                int success = 0;
                int noRoute = 0;
                int noReason = 0;
                int noResource = 0;
                int wrongRoute = 0;

                //遍历该sheet的行
                for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    XSSFRow row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }
                    //筛选主叫终端号
                    int cloumnFrom = 2;
                    XSSFCell callAddressFrom = row.getCell(cloumnFrom);

                    //筛选被叫终端号
                    int cloumnTo = 4;
                    XSSFCell callAddressTo = row.getCell(cloumnTo);

                    //筛选结束原因
                    int columnReason = 6;
                    XSSFCell result = row.getCell(columnReason);
                    System.out.println(result);

                    if (rowNum > 0 && (("成功".equals(result.toString())) || ("无目的路由".equals(result.toString())) || ("未指定的原因".equals(result.toString())) || ("交换路由错误".equals(result.toString())) || ("资源未指定".equals(result.toString())))) {
                        userAll.add(row.getCell(2).toString());
                        userAll.add(row.getCell(4).toString());
                        if ("成功".equals(result.toString())) {
                            System.out.println(111);
                            success++;
                        } else if ("无目的路由".equals(result.toString())) {
                            noRoute++;
                        } else if ("未指定的原因".equals(result.toString())) {
                            noReason++;
                        } else if ("交换路由错误".equals(result.toString())) {
                            wrongRoute++;
                        } else if ("资源未指定".equals(result.toString())) {
                            noResource++;
                        }

                    }

                    try (FileOutputStream outputStream = new FileOutputStream(file);
                         OutputStreamWriter streamWriter = new OutputStreamWriter(outputStream)) {

                        streamWriter.append("========================" + description + "========================\n");
                        streamWriter.append(">>>>>>>>系统话音数据<<<<<<<<\n");

                        streamWriter.append("总呼叫数：" + allRecordLines + "\n");
                        streamWriter.append("呼叫成功数：" + success + "\n");
                        float a = (float) allRecordLines;
                        float b = (float) success;
                        streamWriter.append("呼通率：" + b * 100 / a + "%\n");
                        streamWriter.append("无目的路由：" + noRoute + "\n");
                        streamWriter.append("未指定的原因：" + noReason + "\n");
                        streamWriter.append("资源未指定：" + noResource + "\n");
                        streamWriter.append("交换路由错误：" + wrongRoute + "\n");
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
}