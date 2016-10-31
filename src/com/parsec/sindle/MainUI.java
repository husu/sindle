package com.parsec.sindle;

import com.parsec.sindle.model.MarketData;
import com.parsec.sindle.model.XlsData;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @auther:husu
 * @version:1.16
 * @date 16/10/10.
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
        JFrame frame = new JFrame("Sindle DBD  公1.16 统计版");
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
                    log.append(e1.toString()+"\n");
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
        } else {
             if (e.getSource() == splitExcelButton) {

                 try {
                     validateNumber();
                 } catch (Exception e1) {
                     log.append(e1.getMessage() + "\n");
                     return;
                 }

                 if (outString == null) {
                     log.append("请选择要计算的Excel表" + "\n");
                 } else {
                     new Thread(() -> {     //我觉得这么做不科学，但是。。。你想怎样
                         log.append("正在载入Excel表，么么哒" + "\n");
                         ExcelReader excelReader = new ExcelReader();
                         try {
                             int from = Integer.parseInt(this.maField.getText());
                             int to = Integer.parseInt(this.a60TextField.getText());
                             int tableNum = to - from  +1  ;
                             int pageSize = 30;
                             int blockNum = tableNum / pageSize;

                             blockNum =  tableNum%pageSize>0?blockNum+1:blockNum;

                             if(blockNum>1){
                                 log.append("由于表的数量太多，将生成" + blockNum + "个文件\n");
                             }
                             for(int n =0;n<blockNum;n++){
                                 if(blockNum>1){
                                     log.append("正在生成第"+(n+1)+"个文件\n");
                                 }
                                 int curFrom = from + pageSize*n;
                                 int curTo = n==blockNum-1?to:(pageSize+curFrom-1);
                                 if ((curTo - curFrom + 1) > 9) {
                                     log.append("正在复制并生成MA的中间数据，共有" + (curTo - curFrom + 1) + "个表要生成，可想而知会很慢\n");
                                 } else {
                                     log.append("正在复制并生成MA的中间数据，共有" + (curTo - curFrom + 1) + "个表要生成\n");
                                 }
                                 File maFile = excelReader.createMADataSheet(fc.getSelectedFile(), curFrom, curTo);
                                 log.append("生成MA中间数据完毕\n");
                                 log.append("开始计算交易数据");

                                 XlsData xlsData;
                                 List<Map<String, String>> statList = new ArrayList<>();

                                 Map<String, String> statMap = null;
                                 for (int i = curFrom; i <= curTo; i++) {

                                     xlsData = excelReader.loadXls(maFile, i);

                                     log.append("开始分析计算MA" + i + "数据，么么哒\n");

                                     List<MarketData> tradeList = excelReader.analyseData(xlsData.getMdList(), xlsData.getStopLossLine(), i);

                                     //总共交易次数
                                     statMap = new HashMap<String, String>() {{
                                         put("tradeCount", tradeList.size() + "");
                                     }};

                                     statMap.put("sumWithStop", String.format("Sum('MA%d汇总'!N2:N%d)",i,tradeList.size()+1));  //有止损汇总
                                     statMap.put("sumWithoutStop",String.format("Sum('MA%d汇总'!$M$2:$M$%d)",i,tradeList.size()+1,xlsData.getStopLossLine())); //无止损汇总
                                     statMap.put("winCountWithStop",String.format("COUNTIF('MA%d汇总'!$N$2:$N$%d,\">0\")",i,tradeList.size()+1)); //有止损盈利次数
                                     statMap.put("lossCountWithStop",String.format("COUNTIF('MA%d汇总'!$N$2:$N$%d,\"<0\")",i,tradeList.size()+1)); //有止损亏损次数
                                     statMap.put("maxWithStop",String.format("MAX('MA%d汇总'!$N$2:$N$%d)",i,tradeList.size()+1)); //单次最大盈利
                                     statMap.put("minWithStop",String.format("MIN('MA%d汇总'!$N$2:$N$%d)",i,tradeList.size()+1));//单次最大亏损
                                     statMap.put("winCountWithoutStop",String.format("COUNTIF('MA%d汇总'!$M$2:$M$%d,\">0\")",i,tradeList.size()+1));//无止损盈利次数
                                     statMap.put("lossCountWithoutStop",String.format("COUNTIF('MA%d汇总'!$M$2:$M$%d,\"<0\")",i,tradeList.size()+1));//无止损亏损次数
                                     statMap.put("maxWithoutStop",String.format("MAX('MA%d汇总'!$M$2:$M$%d)",i,tradeList.size()+1));//单次最大盈利
                                     statMap.put("minWithoutStop",String.format("MIN('MA%d汇总'!$M$2:$M$%d)",i,tradeList.size()+1));//单次最大亏损


                                     statList.add(statMap);//将统计的Map加入到统计的列表
                                     excelReader.modify(maFile, tradeList, i);


                                     log.append("MA" + i + "处理完毕，亚克西\n");
                                 }


                                 log.append("正在处理汇总表");

                                 excelReader.fullStatTable(maFile,statList,from);

                                 log.append("处理汇总表结束\n");

                                 log.append("文件写入完毕，文件路径为 ：" + maFile.getPath() + "\n");
                             }



                         } catch (Exception e1) {
                             log.append("错误：" + e1.getMessage());
                             e1.printStackTrace();
                         }
                     }).start();

                 }
             }
         }
    }

}
