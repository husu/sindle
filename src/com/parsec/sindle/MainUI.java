package com.parsec.sindle;

import com.parsec.sindle.model.MarketData;
import com.parsec.sindle.model.XlsData;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.net.URISyntaxException;


import java.util.List;

/**
 * @auther:husu
 * @version:1.0
 * @date 15/6/6.
 */
public class MainUI implements ActionListener {

    private JPanel panel1;
    private JTextArea log;
    private JButton chooseFileButton;
    private JButton splitExcelButton;
    private JPanel panel2;
    private JPanel logPanel;
    private String saveFileName;

    static private final String newline = "\n";
    JFileChooser fc;
    String outString, inputFile3;

    public static void main(String[] args) {


        SwingUtilities.invokeLater(() -> {
            //Turn off metal's use of bold fonts
            UIManager.put("swing.boldMetal", Boolean.FALSE);
            try {
                createAndShowGUI();
            } catch (URISyntaxException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        });
    }

    private static void createAndShowGUI() throws URISyntaxException{
        JFrame frame = new JFrame("Sindle DBD 1.02");
        frame.setContentPane(new MainUI().panel1);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);



    }

    public MainUI() {
        chooseFileButton.addActionListener(this);
        splitExcelButton.addActionListener(this);
        fc = new JFileChooser();
    }

    private void createUIComponents() {
        // TODO: place custom component creation code here
    }

    public void actionPerformed(ActionEvent e) {
         if (e.getSource() == chooseFileButton) {
            if (outString != null) {
                fc.setCurrentDirectory(new File(outString));
            } else {
                File startFile = null;
                try {
                    startFile = new File(MainUI.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath());
                } catch (URISyntaxException e1) {
                    // TODO Auto-generated catch block
                    e1.printStackTrace();
                    log.append(e1.toString());
                }
                fc.setCurrentDirectory(startFile);
            }

            int returnVal = fc.showOpenDialog(panel1);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fc.getSelectedFile();
                inputFile3 = file.getPath();
                outString = fc.getSelectedFile().getPath();
                log.append("Opening origin file: " + file.getPath() + newline);
            } else {
                log.append("Open command cancelled by user." + newline);
            }
            log.setCaretPosition(log.getDocument().getLength());
        } else if (e.getSource() == splitExcelButton) {
            if (outString == null) {
                log.append("请选择要计算的Excel表" + newline);
            }else {
               new Thread(() -> {     //我觉得这么做不科学，但是。。。你想怎样
                   log.append("正在载入Excel表，么么哒" + newline);


                   ExcelReader excelReader = new ExcelReader();

                   try {
                       XlsData xlsData =  excelReader.loadXls(fc.getSelectedFile());

                       log.append("载入成功，么么哒" + newline);
                       log.append("分析计算中，么么哒" + newline);

                       List<MarketData> tradeList = excelReader.analyseData(xlsData.getMdList(),xlsData.getStopLossLine());

                       log.append("分析计算完毕，么么哒" + newline);
                       log.append("复制文件，么么哒" + newline);

                       File newFile = excelReader.pasteFile(fc.getSelectedFile());

                       log.append("文件写入中，么么哒" + newline);
                       excelReader.modify(newFile,tradeList);

                       log.append("文件写入完毕，文件路径为 ：" + newFile.getPath() + newline);

                   } catch (Exception e1) {
                       log.append("错误：" + e1.getMessage());
                       e1.printStackTrace();
                   }
               }).start();

            }
        }
    }

}
