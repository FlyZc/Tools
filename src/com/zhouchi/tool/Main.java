package com.zhouchi.tool;

import java.awt.Container;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import javax.swing.*;

/**
 * @author zhouchi
 */
public class Main extends JFrame {
    public Main() {
        Container cont = getContentPane();

        //此处要使用绝对布局，取消布局管理器
        setLayout(null);

        //显示统计信息
        final JTextArea jt = new JTextArea();
        jt.setBounds(5, 45, 385, 540);
        JScrollPane scroll = new JScrollPane();
        scroll.setBounds(5, 45, 385, 540);

        //把定义的JTextArea放到JScrollPane里面去
        scroll.setViewportView(jt);
        scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);

        //显示提示信息
        final JTextArea jtTt = new JTextArea();
        jtTt.setBounds(405, 45, 385, 540);
        JScrollPane scrollInfo = new JScrollPane();
        scrollInfo.setBounds(405, 45, 385, 540);

        //把定义的JTextArea放到JScrollPane里面去
        scrollInfo.setViewportView(jtTt);
        scrollInfo.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);


        final JButton jb = new JButton("二代网系");
        jb.setBounds(5, 5, 385, 40);
        jb.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {

                   /* FDMA.main(null);
                    TD30.main(null);
                    TD31.main(null);
                    TD31.main(null);
                    TD32.main(null);
                    TD34.main(null);
                    KGR.main(null);*/

                    print(jt, "E:\\Daily\\FDMA.txt");
                    print(jt, "E:\\Daily\\FDMADH.txt");
                    print(jt, "E:\\Daily\\FDMANH.txt");
                    print(jt, "E:\\Daily\\FDMATH.txt");
                    print(jt, "E:\\Daily\\FDZM.txt");
                    print(jt, "E:\\Daily\\FDZYZDGZ.txt");
                    print(jt, "E:\\Daily\\FDZYZB.txt");
                    print(jt, "E:\\Daily\\FDJLW.txt");
                    print(jt, "E:\\Daily\\TD30.txt");
                    print(jt, "E:\\Daily\\TDDH30.txt");
                    print(jt, "E:\\Daily\\TDNH30.txt");
                    print(jt, "E:\\Daily\\TDZYZDGZ30.txt");
                    print(jt, "E:\\Daily\\TD31.txt");
                    print(jt, "E:\\Daily\\TDTH31.txt");
                    print(jt, "E:\\Daily\\TDNH31.txt");
                    print(jt, "E:\\Daily\\TD32.txt");
                    print(jt, "E:\\Daily\\TDDH32.txt");
                    print(jt, "E:\\Daily\\TDTH32.txt");
                    print(jt, "E:\\Daily\\TDNH32.txt");
                    print(jt, "E:\\Daily\\TDZYZDGZ32.txt");
                    print(jt, "E:\\Daily\\TD34.txt");
                    print(jt, "E:\\Daily\\TDDH34.txt");
                    print(jt, "E:\\Daily\\TDTH34.txt");
                    print(jt, "E:\\Daily\\TDNH34.txt");
                    print(jt, "E:\\Daily\\KGR.txt");
                    print(jt, "E:\\Daily\\KGRDH.txt");
                    print(jt, "E:\\Daily\\KGRTH.txt");
                    print(jt, "E:\\Daily\\KGRNH.txt");

                } catch (Exception ex) {

                    ex.printStackTrace();
                }
            }
        });

        final JButton jbTt = new JButton("TT系统");
        jbTt.setBounds(405, 5, 385, 40);
        jbTt.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    TT.main(null);
                    print(jtTt, "E:\\Daily\\TT1.txt");
                    print(jtTt, "E:\\Daily\\TT2.txt");

                } catch (Exception ex) {

                    ex.printStackTrace();
                }
            }
        });

        cont.add(jb);
        cont.add(jbTt);
        cont.add(scroll);
        cont.add(scrollInfo);

        setBounds(300,30,830,600);
        setVisible(true);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
    }
    public static void print(JTextArea jt, String pathname) {
        try (FileReader reader = new FileReader(pathname);
             BufferedReader br = new BufferedReader(reader)
        ) {
            String line;
            while ((line = br.readLine()) != null) {
                // 一次读入一行数据
                jt.setText(jt.getText() + line + "\n");
            }
        } catch (IOException ee) {
            ee.printStackTrace();
        }
    }
    public static void main(String[] args) {
        new Main();
    }
}