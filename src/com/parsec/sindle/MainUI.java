package com.parsec.sindle;

import com.parsec.sindle.model.MarketData;
import com.parsec.sindle.model.XlsData;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.net.URISyntaxException;


import java.util.List;
import java.util.stream.Stream;

/**
 * @auther:husu
 * @version:1.04
 * @date 15/6/6.
 */
public class MainUI implements ActionListener {

    private JPanel panel1;
    private JTextArea log;
    private JButton chooseFileButton;
    private JButton splitExcelButton;
    private JTextField maField;
    private JTextField a60TextField;

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

    private void validateNumber() throws Exception{
        Integer maFrom = Integer.parseInt(maField.getText());
        Integer maTo = Integer.parseInt(a60TextField.getText());
        if(maFrom>maTo){
            throw new Exception("MA范围开始数值不得小于结束数值");
        }

    }

    private static void createAndShowGUI() throws URISyntaxException{
        JFrame frame = new JFrame("Sindle DBD 1.04");
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
                log.append("Opening origin file: " + file.getPath() + "\n");
            } else {
                log.append("Open command cancelled by user." + "\n");
            }
            log.setCaretPosition(log.getDocument().getLength());
        } else if (e.getSource() == splitExcelButton) {

             try {
                 validateNumber();
             } catch (Exception e1) {
                 log.append(e1.getMessage() + "\n");
                 return;
             }

             if (outString == null) {
                log.append("请选择要计算的Excel表" + "\n");
            }else {
               new Thread(() -> {     //我觉得这么做不科学，但是。。。你想怎样
                   log.append("正在载入Excel表，么么哒" + "\n");
                   ExcelReader excelReader = new ExcelReader();

                   try {

                       int from =Integer.parseInt(this.maField.getText());
                       int to = Integer.parseInt(this.a60TextField.getText());


                       log.append("正在复制并生成MA的中间数据，共有" + (to-from+1) + "个表要生成，可想而知会很慢\n");
                       File maFile = excelReader.createMADataSheet(fc.getSelectedFile(),from,to);


                       log.append("生成MA中间数据完毕\n");


                       log.append("开始计算交易数据");



                       XlsData  xlsData;
                       for(int i=from;i<=to;i++){

                           xlsData =  excelReader.loadXls(maFile,i);

                           log.append("开始分析计算MA" + i + "数据，么么哒\n");

                           List<MarketData> tradeList = excelReader.analyseData(xlsData.getMdList(),xlsData.getStopLossLine());


                           excelReader.modify(maFile,tradeList,i);

                           log.append("MA" + i + "处理外币，么么哒\n");


                       }

                       log.append("文件写入完毕，文件路径为 ：" + maFile.getPath() + "\n");


                   } catch (Exception e1) {
                       log.append("错误：" + e1.getMessage());
                       e1.printStackTrace();
                   }
               }).start();

            }
        }
    }

}
