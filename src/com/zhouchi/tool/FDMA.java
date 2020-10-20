package com.zhouchi.tool;

/**
 * @Project: Tools
 * @Description: 筛选FDMA数据
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

public class FDMA {
    public static void main(String[] args) throws Exception {
        File fileInfo = new File("E:\\Daily\\Information.txt");
        FileOutputStream outputStreamInfo = new FileOutputStream(fileInfo);
        OutputStreamWriter streamWriterInfo = new OutputStreamWriter(outputStreamInfo);

        //读取需统计用户的数据
        InputStream streamUser = new FileInputStream("E:\\Daily\\FDMA.xlsx");
        if (streamUser == null) {
            streamWriterInfo.append("FDMA重保用户名单<User.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUser = new XSSFWorkbook(streamUser);
        //获取excel表的第一个sheet
        XSSFSheet sheetUser = wbUser.getSheetAt(0);
        if (sheetUser == null) {
            return;
        }

        //读取FD DH方向用户的数据
        InputStream streamUserDH = new FileInputStream("E:\\Daily\\FDDH.xlsx");
        if (streamUserDH == null) {
            streamWriterInfo.append("FDMA DH方向用户名单<FDDH.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserDH = new XSSFWorkbook(streamUserDH);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserDH = wbUserDH.getSheetAt(0);
        if (sheetUserDH == null) {
            return;
        }

        //读取FD NH方向用户的数据
        InputStream streamUserNH = new FileInputStream("E:\\Daily\\FDNH.xlsx");
        if (streamUserNH == null) {
            streamWriterInfo.append("FDMA NH方向用户名单<FDNH.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserNH = new XSSFWorkbook(streamUserNH);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserNH = wbUserNH.getSheetAt(0);
        if (sheetUserNH == null) {
            return;
        }

        //读取FD TH方向用户的数据
        InputStream streamUserTH = new FileInputStream("E:\\Daily\\FDTH.xlsx");
        if (streamUserTH == null) {
            streamWriterInfo.append("FDMA NH方向用户名单<FDTH.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserTH = new XSSFWorkbook(streamUserTH);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserTH = wbUserTH.getSheetAt(0);
        if (sheetUserTH == null) {
            return;
        }

        //读取FD ZYZDGZ方向用户的数据
        InputStream streamUserZYZDGZ = new FileInputStream("E:\\Daily\\FDZYZDGZ.xlsx");
        if (streamUserZYZDGZ == null) {
            streamWriterInfo.append("FDMA ZYZDGZ方向用户名单<FDZYZDGZ.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserZYZDGZ = new XSSFWorkbook(streamUserZYZDGZ);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserZYZDGZ = wbUserZYZDGZ.getSheetAt(0);
        if (sheetUserZYZDGZ == null) {
            return;
        }

        //读取FD ZYZB方向用户的数据
        InputStream streamUserZYZB = new FileInputStream("E:\\Daily\\FDZYZB.xlsx");
        if (streamUserZYZB == null) {
            streamWriterInfo.append("FDMA ZYZB方向用户名单<FDZYZB.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserZYZB = new XSSFWorkbook(streamUserZYZB);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserZYZB = wbUserZYZB.getSheetAt(0);
        if (sheetUserZYZB == null) {
            return;
        }

        //读取FD ZYZB方向用户的数据
        InputStream streamUserJLW = new FileInputStream("E:\\Daily\\FDJLW.xlsx");
        if (streamUserJLW == null) {
            streamWriterInfo.append("FDMA JLW方向用户名单<FDJLW.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserJLW = new XSSFWorkbook(streamUserJLW);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserJLW = wbUserJLW.getSheetAt(0);
        if (sheetUserJLW == null) {
            return;
        }

        //读取FD ZM方向用户的数据
        InputStream streamUserZM = new FileInputStream("E:\\Daily\\FDZM.xlsx");
        if (streamUserZM == null) {
            streamWriterInfo.append("FDMA ZM方向用户名单<FDZM.xlsx>不存在或文件命名错误，请核实！！！\n");
            return;
        }
        XSSFWorkbook wbUserZM = new XSSFWorkbook(streamUserZM);
        //获取excel表的第一个sheet
        XSSFSheet sheetUserZM = wbUserZM.getSheetAt(0);
        if (sheetUserZM == null) {
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

        Set<String> userZYZDGZ = new HashSet<String>(sheetUserZYZDGZ.getLastRowNum());
        Set<String> userZYZDGZResult = new HashSet<String>(sheetUserZYZDGZ.getLastRowNum());

        Set<String> userZYZB = new HashSet<String>(sheetUserZYZB.getLastRowNum());
        Set<String> userZYZBResult = new HashSet<String>(sheetUserZYZB.getLastRowNum());

        Set<String> userJLW = new HashSet<String>(sheetUserJLW.getLastRowNum());
        Set<String> userJLWResult = new HashSet<String>(sheetUserJLW.getLastRowNum());

        Set<String> userZM = new HashSet<String>(sheetUserZM.getLastRowNum());
        Set<String> userZMResult = new HashSet<String>(sheetUserZM.getLastRowNum());

        File fileAll = new File("E:\\Daily\\FDMA.txt");
        File fileDH = new File("E:\\Daily\\FDMADH.txt");
        File fileTH = new File("E:\\Daily\\FDMATH.txt");
        File fileNH = new File("E:\\Daily\\FDMANH.txt");
        File fileZYZDGZ = new File("E:\\Daily\\FDZYZDGZ.txt");
        File fileZYZB = new File("E:\\Daily\\FDZYZB.txt");
        File fileJLW = new File("E:\\Daily\\FDJLW.txt");
        File fileZM = new File("E:\\Daily\\FDZM.txt");

        //获取ZB用户站列表
        readUserSet(sheetUserDH, userDH);
        readUserSet(sheetUserNH, userNH);
        readUserSet(sheetUserTH, userTH);
        readUserSet(sheetUserZYZDGZ, userZYZDGZ);
        readUserSet(sheetUserZYZB, userZYZB);
        readUserSet(sheetUserJLW, userJLW);
        readUserSet(sheetUserZM, userZM);

        //获取所有通信用户站列表;
        String path = "E:\\Daily\\FDMA.xlsx";

        calAllRate(path, userAll,fileAll,"FDMA全网呼通率统计");
        calIRate(path, userAll, userDH, userDHResult, fileDH,"FDMADH方向呼通率统计");
        calIRate(path, userAll, userNH, userNHResult, fileNH,"FDMANH方向呼通率统计");
        calIRate(path, userAll, userTH, userTHResult, fileTH,"FDMATH方向呼通率统计");
        calIRate(path, userAll, userZM, userZMResult, fileZM, "FDMAZM方向呼通率统计");
        calIRate(path, userAll, userZYZDGZ, userZYZDGZResult, fileZYZDGZ, "FDZYZDGZ呼通率统计");
        calIRate(path, userAll, userZYZB, userZYZBResult, fileZYZB, "FDZYZB呼通率统计");
        calIRate(path, userAll, userJLW, userJLWResult, fileJLW, "FDJLW呼通率统计");
    }

    public static String readCellMethod(XSSFCell cell) {
        return cell.getStringCellValue();
    }

    public static void calIRate(String allPath, Set<String> userAll, Set<String> user, Set<String> userResult, File file, String description) {
        try {
            //读取所有通话记录的信息
            InputStream stream = new FileInputStream(allPath);
            if (stream == null) {
                return;
            } else {
                XSSFWorkbook wb = new XSSFWorkbook(stream);

                //获取excel表的第一个sheet
                XSSFSheet sheet = wb.getSheetAt(0);

                if (sheet == null) {
                    return;
                }

                //用来记录筛选出的用户站通话成功的数量
                int successCount = 0;
                //用来记录筛选出的用户站等待...的数量
                int waitCount = 0;
                //用来记录筛选出的用户站KDE异常的数量
                int kdcErrorCount = 0;
                //用来记录筛选出的用户站所有拨打电话的数量
                int allUserCount = 0;

                //遍历该sheet的行
                for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    XSSFRow row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }

                    //筛选主叫地址
                    int cloumn1 = 1;
                    XSSFCell callAddress = row.getCell(cloumn1);

                    //筛选呼叫失败原因
                    int column = 8;
                    XSSFCell reason = row.getCell(column);

                    //呼叫结果
                    int column2 = 7;
                    XSSFCell result = row.getCell(column2);

                    if (callAddress == null) {
                        continue;
                    } else {
                        if (user.contains(callAddress.toString())) {
                            userResult.add(callAddress.toString());
                            System.out.println("FDFD " + callAddress.toString());
                            if (("成功".equals(result.toString()))) {
                                successCount++;
                            } else if ("KDM异常".equals(reason.toString()) || "等待通信检测应答或者干扰检测上报超时".equals(reason.toString())) {
                                if ("KDM异常".equals(reason.toString())) {
                                    kdcErrorCount++;
                                } else if ("等待通信检测应答或者干扰检测上报超时".equals(reason.toString())) {
                                    waitCount++;
                                }
                            }
                        }
                    }
                    allUserCount = successCount + kdcErrorCount + waitCount;
                    try (FileOutputStream outputStream = new FileOutputStream(file);
                         OutputStreamWriter streamWriter = new OutputStreamWriter(outputStream)) {
                        streamWriter.append("========================" + description + "========================\n");

                        streamWriter.append("\n********重保********\n");
                        streamWriter.append("总呼叫数：" + allUserCount + "\n");
                        streamWriter.append("呼叫成功数：" + successCount + "\n");
                        float a = (float) allUserCount;
                        float b = (float) successCount;
                        streamWriter.append("呼通率：" + b * 100 / a + "%\n");
                        streamWriter.append("KDM异常：" + kdcErrorCount + "\n");
                        streamWriter.append("等待通信检测应答或者干扰检测上报超时：" + waitCount + "\n");
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

                //用来记录筛选出的用户站通话成功的数量
                int successCountAll = 0;
                //用来记录筛选出的用户站等待...的数量
                int waitCountAll = 0;
                //用来记录筛选出的用户站KDE异常的数量
                int kdcErrorCountAll = 0;
                //用来记录用户站所有拨打电话的数量
                int countAll = 0;

                int allRecordLines = sheet.getLastRowNum() - 1;

                //遍历该sheet的行
                for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    XSSFRow row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }
                    if (rowNum > 1) {
                        userAll.add(row.getCell(1).toString());
                        userAll.add(row.getCell(3).toString());
                    }

                    //筛选主叫地址
                    int cloumn1 = 1;
                    XSSFCell callAddress = row.getCell(cloumn1);

                    //筛选呼叫失败原因
                    int column = 8;
                    XSSFCell reason = row.getCell(column);

                    //呼叫结果
                    int column2 = 7;
                    XSSFCell result = row.getCell(column2);

                    if (callAddress == null) {
                        continue;
                    } else {
                        if ("成功".equals(result.toString())) {
                            successCountAll++;
                        } else if ("KDM异常".equals(reason.toString()) || "等待通信检测应答或者干扰检测上报超时".equals(reason.toString())) {
                            if ("KDM异常".equals(reason.toString())) {
                                kdcErrorCountAll++;
                            } else if ("等待通信检测应答或者干扰检测上报超时".equals(reason.toString())) {
                                waitCountAll++;
                            }
                        }
                    }

                    countAll = successCountAll + kdcErrorCountAll + waitCountAll;
                    float all = (float) countAll;
                    float success = (float) successCountAll;

                    try (FileOutputStream outputStream = new FileOutputStream(file);
                             OutputStreamWriter streamWriter = new OutputStreamWriter(outputStream)) {

                        streamWriter.append("========================" + description + "========================\n");
                        streamWriter.append(">>>>>>>>全网<<<<<<<<\n");
                        streamWriter.append("总呼叫数：" + countAll + "\n");
                        streamWriter.append("呼叫成功数：" + successCountAll + "\n");
                        streamWriter.append("呼通率：" + success * 100 / all + "%\n");
                        streamWriter.append("KDM异常：" + kdcErrorCountAll + "\n");
                        streamWriter.append("等待通信检测应答或者干扰检测上报超时：" + waitCountAll + "\n");
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
                System.out.println("FD" + String.valueOf(row.getCell(0).getStringCellValue()));

            }
        }
    }
}