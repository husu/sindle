package com.parsec.sindle;

import com.parsec.sindle.model.MarketData;
import com.parsec.sindle.model.TradeType;
import com.parsec.sindle.model.XlsData;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @auther:husu
 * @version:1.0
 * @date 16/9/6.
 */
public class ExcelReader {


    public ExcelReader(Double slp, Double upSLP, Double swp) {
        this.slp = slp;
        this.upSLP = upSLP;
        this.swp = swp;
    }

    private static final String STAT_TABLE  = "交易汇总";    //汇总表的名字


    private Double slp=0.0;//止损点
    private Double upSLP=0.0;//上浮止损点
    private Double swp=0.0;//止盈点



    /**
     * 写入统计表数据
     * @param file
     * @param statList
     */
    public void fullStatTable(File file,List<Map<String,String>> statList,int from) throws Exception {
        InputStream is = new FileInputStream(file);


        String fileName = file.getName();

        Workbook wbs = this.getWorkbookInstance(fileName,is);

        Sheet statSheet = wbs.getSheet(STAT_TABLE);

        int n ;
        Map<String,String> formulaMap;
        Row curRow;
        for(int i=0;i<statList.size();i++){
            formulaMap = statList.get(i);
            n =0;
            curRow = statSheet.getRow(i+2);
            if(curRow==null){
                curRow = statSheet.createRow(i+2);
            }
            getEditingCell(curRow,n++).setCellValue("MA" + (from+i));
            getEditingCell(curRow,n++).setCellValue(formulaMap.get("tradeCount"));  //交易次数
            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("sumWithStop"));
            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("sumWithoutStop"));

            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("winCountWithStop"));
            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("lossCountWithStop"));
            getEditingCell(curRow,n++).setCellFormula("E" + (i+3) + "/B" + (i+3)); //盈利次数比例
            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("maxWithStop"));
            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("minWithStop"));

            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("winCountWithoutStop"));
            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("lossCountWithoutStop"));
            getEditingCell(curRow,n++).setCellFormula("J" + (i+3) + "/B" +(i+3));

            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("maxWithoutStop"));
            getEditingCell(curRow,n++).setCellFormula(formulaMap.get("minWithoutStop"));


        }

        this.writeWbs(file,wbs);
    }


    /**
     * 创建含有MA数据的各个表
     * @param file
     * @return
     */
    public File createMADataSheet(File file,int from,int to) throws Exception {

        StringBuffer newFileName = new StringBuffer(file.getName().substring(0,file.getName().lastIndexOf(".")))
                .append("-MADATA-").append(new Date().getTime()).append(file.getName().substring(file.getName().lastIndexOf(".")));


        File mafile = this.pasteFile(file,newFileName.toString());


        InputStream is = new FileInputStream(mafile);
        org.apache.poi.ss.usermodel.Workbook wbs;

        wbs = this.getWorkbookInstance(mafile.getName(),is);

        Sheet maSheet;

        Sheet firstSheet = wbs.getSheetAt(0);


        //先创建一个汇总表

        Sheet statSheet =   wbs.createSheet(STAT_TABLE);
        Row titleRow = statSheet.getRow(0);
        if(titleRow==null){
            titleRow =  statSheet.createRow(0);
        }


        getEditingCell(titleRow,1).setCellValue("交易汇总");
        getEditingCell(titleRow,2).setCellValue("交易汇总");
        getEditingCell(titleRow,3).setCellValue("交易汇总");
        getEditingCell(titleRow,4).setCellValue("有止损统计");
        getEditingCell(titleRow,5).setCellValue("有止损统计");
        getEditingCell(titleRow,6).setCellValue("有止损统计");
        getEditingCell(titleRow,7).setCellValue("有止损统计");
        getEditingCell(titleRow,8).setCellValue("有止损统计");
        getEditingCell(titleRow,9).setCellValue("无止损统计");
        getEditingCell(titleRow,10).setCellValue("无止损统计");
        getEditingCell(titleRow,11).setCellValue("无止损统计");
        getEditingCell(titleRow,12).setCellValue("无止损统计");
        getEditingCell(titleRow,13).setCellValue("无止损统计");


        titleRow = statSheet.getRow(1);
        if(titleRow==null){
            titleRow =  statSheet.createRow(1);
        }

        getEditingCell(titleRow,1).setCellValue("交易次数");
        getEditingCell(titleRow,2).setCellValue("有止损汇总");
        getEditingCell(titleRow,3).setCellValue("无止损汇总");
        getEditingCell(titleRow,4).setCellValue("有止损盈利次数");
        getEditingCell(titleRow,5).setCellValue("有止损亏损次数");
        getEditingCell(titleRow,6).setCellValue("盈利次数比例");
        getEditingCell(titleRow,7).setCellValue("单次最大盈利");
        getEditingCell(titleRow,8).setCellValue("单次最大亏损");
        getEditingCell(titleRow,9).setCellValue("无止损盈利次数");
        getEditingCell(titleRow,10).setCellValue("无止损亏损次数");
        getEditingCell(titleRow,11).setCellValue("盈利次数比例");
        getEditingCell(titleRow,12).setCellValue("单次最大盈利");
        getEditingCell(titleRow,13).setCellValue("单次最大亏损");





         // 复制多个MA表

        Row r1;
        Row stopRow; //止损所在的行
        for(int i=from;i<=to;i++){

            maSheet = wbs.createSheet("MA" + i);

            stopRow =  maSheet.getRow(2);
            if(stopRow==null){
                stopRow = maSheet.createRow(2);

            }

            getEditingCell(stopRow,0).setCellValue("止损");
            getEditingCell(stopRow,1).setCellValue(firstSheet.getRow(2).getCell(1).getNumericCellValue());


            StringBuffer formula ;
            for(int r=4;r<=firstSheet.getLastRowNum();r++){      //遍历行

                r1 = maSheet.getRow(r);
                if(r1==null){
                    r1= maSheet.createRow(r);
                }
                for(int c=0;c<22;c++){ //填充列
                    if(getEditingCell(firstSheet.getRow(r),c).getCellType()==Cell.CELL_TYPE_STRING) {
                        getEditingCell(r1, c).setCellValue(firstSheet.getRow(r).getCell(c).getStringCellValue());
                    }else if(getEditingCell(firstSheet.getRow(r),c).getCellType()==Cell.CELL_TYPE_NUMERIC) {
                        getEditingCell(r1, c).setCellValue(firstSheet.getRow(r).getCell(c).getNumericCellValue());
                    }else if(getEditingCell(firstSheet.getRow(r),c).getCellType()==Cell.CELL_TYPE_FORMULA){
                        getEditingCell(r1, c).setCellFormula(firstSheet.getRow(r).getCell(c).getCellFormula());
                    }else{
                        getEditingCell(r1, c).setCellValue("");
                    }
                }

                if(r-5>i-2){ //跳过前 i-1行
                    formula = new StringBuffer("average(E").append(r-i+2).append(":E").append(r+1).append(")");
                    getEditingCell(r1,6).setCellFormula(formula.toString());//MA30
                    getEditingCell(r1,7).setCellFormula("(D" + (r + 1) + ">G" + (r + 1) + ")*1");//最低计算
                    getEditingCell(r1,8).setCellFormula("(C" + (r + 1) + ">G" + (r + 1) + ")*1"); //最高计算
                    getEditingCell(r1,9).setCellFormula("IF(H" + r + "=I" + r + ",I" + r + ",J" + r + ")"); //多空计算

                }
            }


        }


        //写入文件

        this.writeWbs(mafile,wbs);
        wbs =null;
        return mafile;
    }


    /**
     * 从文件中读取行情数据<br>
     * @param f
     * @return
     * @throws IOException
     */
    public XlsData loadXls(File f,Integer maNum) throws Exception {

        InputStream is = new FileInputStream(f);
        org.apache.poi.ss.usermodel.Workbook wbs;

        String fileName = f.getName();

        wbs = this.getWorkbookInstance(fileName,is);

        org.apache.poi.ss.usermodel.Sheet childSheet = wbs.getSheet("MA"+maNum);


        Double mostLossLine = checkNull(java.lang.Math.abs(childSheet.getRow(2).getCell(1).getNumericCellValue())*(-1),"止损设置，B3单元格");


        MarketData marketData;
        Double preDk = null;   //记录上一个多空状态
        Row curRow;

        childSheet.setForceFormulaRecalculation(true);


        List<MarketData> mdList= new ArrayList<>();

        FormulaEvaluator evaluator = this.getFormulaEvalatorInstance(fileName,wbs);

        boolean flag = false;

        for(int i=childSheet.getLastRowNum();i>=(5+maNum-1);i--){
            curRow = childSheet.getRow(i);


            marketData = new MarketData(i);


            flag = false;



            if ((preDk!=null && preDk != this.getFormulaValue(curRow.getCell(9),evaluator)) || i== childSheet.getLastRowNum()){
                flag =true;

            }

            preDk = this.getFormulaValue(curRow.getCell(9),evaluator);



            marketData.setTradePoint(flag);
            marketData.setTradeType(preDk==0.0?TradeType.SHORT:TradeType.LONG);

            marketData.setOpenPrice(checkNull(curRow.getCell(1).getNumericCellValue(),"第" + (i+1) + "行，开盘价"));
            marketData.setHightestPrice(checkNull(curRow.getCell(2).getNumericCellValue(),"第" + (i+1) + "行，最高价"));
            marketData.setLowestPrice(checkNull(curRow.getCell(3).getNumericCellValue(),"第" + (i+1) + "行，最低价"));
            marketData.setClosePrice(checkNull(curRow.getCell(4).getNumericCellValue(),"第" + (i+1) + "行，收盘价"));



            mdList.add(marketData);
        }

        XlsData xlsData = new XlsData(mdList,mostLossLine);


        is.close();

        return xlsData;

    }

    private double getFormulaValue(Cell cell,FormulaEvaluator evaluator){
        if(cell.getCellType()== Cell.CELL_TYPE_FORMULA){
            return evaluator.evaluate(cell).getNumberValue();//获取单元格的值
        }else{
            return cell.getNumericCellValue();
        }
    }

    public List<MarketData> analyseData(List<MarketData> mdList,Double mostLossLine,int maNum){
        List<MarketData> tradeList = mdList.stream().sorted((p1, p2)->(p1.getRowIndex()>p2.getRowIndex()?1:-1)).filter(MarketData::getTradePoint).collect(Collectors.toList());

        tradeList.stream().reduce(new MarketData(4+maNum-1),(p1,p2)-> {  //这个地方的初始值是4，是因为下面做了加1操作
            p2.setPreTradePoint(p1.getRowIndex()+1);
            return p2;
        });


        tradeList.forEach(p->{
            Map<String,String> map =  new HashMap<>();

            //开仓点位
            map.put("openPoint","E" + (mdList.stream().filter(md-> {
                int pp = p.getPreTradePoint()-1;
                pp = pp<(5+maNum-1)?(5+maNum-1):pp;
                return md.getRowIndex() == pp;
            }).max(Comparator.comparing(MarketData::getClosePrice)).get().getRowIndex() + 1));

            double buyPrice = mdList.stream().filter(md-> {
                int pp = p.getPreTradePoint()-1;
                pp = pp<(5+maNum-1)?(5+maNum-1):pp;
                return md.getRowIndex() == pp;
            }).max(Comparator.comparing(MarketData::getClosePrice)).get().getClosePrice();

            p.setBuyPrice(buyPrice);


            //平仓点位
            map.put("sellPoint","E"+(p.getRowIndex()+1));

            //最高价
            map.put("highestPrice","max(C" + (p.getPreTradePoint()+1) + ":C" + (p.getRowIndex()+1) + ")");

            //最低价
//            map.put("lowestPrice",mdList.stream().filter(md-> (md.getRowIndex()>=(p.getPreTradePoint()) && md.getRowIndex()<=p.getRowIndex()))
//                    .min(Comparator.comparing(MarketData::getLowestPrice)).get().getLowestPrice());
            map.put("lowestPrice","min(D"+(p.getPreTradePoint()+1)+":D" + (p.getRowIndex()+1) + ")");


            //结果无止损
//            Double lossNoStop = (p.getTradeType()==TradeType.SHORT)?map.get("openPoint")-map.get("sellPoint"):map.get("sellPoint")-map.get("openPoint");
//            map.put("lossNoStop",lossNoStop);

            String lossNoStop =  (p.getTradeType()==TradeType.SHORT)?("K" + (p.getRowIndex()+1) + "-L" +(p.getRowIndex()+1))
                    :("L" + (p.getRowIndex()+1) + "-K" +  (p.getRowIndex()+1) );
            map.put("lossNoStop",lossNoStop);

            //计算最大亏损
//            Double mostLoss==TradeType.LONG?(map.get("lowestPrice")-map.get("openPoint")):(map.get("openPoint")-map.get("highestPrice")));
            map.put("mostLoss",p.getTradeType()==TradeType.LONG?"P"+(p.getRowIndex()+1)+"-K"+(p.getRowIndex()+1):"K"+(p.getRowIndex()+1) + "-O" + (p.getRowIndex()+1));

            //结果有止损   逻辑是，如果没有止损，则为无止损结果，止损，则为止损值
            map.put("lossStop","IF(R"+(p.getRowIndex()+1)+"<$B$3,$B$3,M" + (p.getRowIndex()+1) + ")");

            //计算最大赢利
//            map.put("mostEarn",p.getTradeType()==TradeType.LONG?(map.get("highestPrice")-map.get("openPoint")):(map.get("openPoint")-map.get("lowestPrice")));
            map.put("mostEarn",p.getTradeType()==TradeType.LONG?"O" + (p.getRowIndex()+1) + "-K" + (p.getRowIndex()+1):"K" +(p.getRowIndex()+1) + "-P" + (p.getRowIndex()+1));

            p.setResultMap(map);
        });

        return tradeList;
    }


    private Double checkNull(Double value,String description) throws Exception{
        if(value==null){
            throw new Exception("妈蛋有个单元格没有填数据这样好吗？位置：" + description);
        }

        if(value==0.0){
            throw new Exception("这个单元格没有填数据或者数据为零，这不科学！位置：" + description);
        }

        return value;
    }


    /**
     * 复制粘贴文件，没有什么新意的老Java代码
     * @param sourceFile 复制文件来源
     * @param newFileName 新文件的名字
     * @return
     * @throws IOException
     */
    public File pasteFile(File sourceFile,String newFileName) throws IOException {

        String sourceFileName = sourceFile.getName();
        String targetFilePath = sourceFile.getParentFile().getPath() + File.separator + sourceFileName.substring(0,sourceFileName.lastIndexOf("."))
              + "-" + (new Date()).getTime() + "(done)"  + "." + sourceFileName.substring(sourceFileName.lastIndexOf(".")+1);

        if(!"".equals(newFileName)){
            targetFilePath =    sourceFile.getParentFile().getPath() + File.separator + newFileName;
        }


        File targetFile = new File(targetFilePath);

        BufferedInputStream inBuff = null;
        BufferedOutputStream outBuff = null;
        try {
            // 新建文件输入流并对它进行缓冲
            inBuff = new BufferedInputStream(new FileInputStream(sourceFile));

            // 新建文件输出流并对它进行缓冲
            outBuff = new BufferedOutputStream(new FileOutputStream(targetFile));

            // 缓冲数组
            byte[] b = new byte[1024 * 5];
            int len;
            while ((len = inBuff.read(b)) != -1) {
                outBuff.write(b, 0, len);
            }
            // 刷新此缓冲的输出流
            outBuff.flush();
        }catch (IOException e){
            throw e;
        }finally {

            // 关闭流
            if (inBuff != null)
                try {
                    inBuff.close();
                } catch (IOException e) {
                    throw e;
                }
            if (outBuff != null)
                try {
                    outBuff.close();
                } catch (IOException e) {
                    throw e;
                }
        }

        return targetFile;
    }

    /**
     * 填表格，好无聊
     * @param targetXlsFile
     * @param tradeList
     */
    public void modify(File targetXlsFile,List<MarketData> tradeList,int maNum) throws Exception {

        InputStream is = new FileInputStream(targetXlsFile);


        String fileName = targetXlsFile.getName();

        Workbook wbs = this.getWorkbookInstance(fileName,is);

        Sheet childSheet = wbs.getSheet("MA" + maNum);



        Sheet newSheet = wbs.createSheet("MA"+maNum+"汇总");
        Row r1= newSheet.createRow(0);
        for(Integer x = 0;x<22;x++){         //复制表头
            getEditingCell(r1,x).setCellValue(childSheet.getRow(4).getCell(x).getStringCellValue());
        }

        CellStyle style =  wbs.createCellStyle();
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(HSSFColor.RED.index);




//        double stopLine = this.getEditingCell(childSheet.getRow(2),1).getNumericCellValue();   //读取止损点
        double stopLine = this.slp.doubleValue();  //读取止损点

        if(stopLine == 0 ){
            throw new Exception("止损未填写");
        }


        tradeList.forEach(p->{          //填充交易点
            Map<String,String> curMap = p.getResultMap();

            Integer i= 10;
            getEditingCell(childSheet.getRow(p.getRowIndex()),i++).setCellFormula(curMap.get("openPoint"));
            getEditingCell(childSheet.getRow(p.getRowIndex()),i++).setCellFormula(curMap.get("sellPoint"));
            getEditingCell(childSheet.getRow(p.getRowIndex()),i++).setCellFormula(curMap.get("lossNoStop"));
            getEditingCell(childSheet.getRow(p.getRowIndex()),i++).setCellFormula(curMap.get("lossStop"));//结果有止损
            getEditingCell(childSheet.getRow(p.getRowIndex()),i++).setCellFormula(curMap.get("highestPrice")); //最高价
            getEditingCell(childSheet.getRow(p.getRowIndex()),i++).setCellFormula(curMap.get("lowestPrice"));
            getEditingCell(childSheet.getRow(p.getRowIndex()),i++).setCellFormula(curMap.get("mostEarn"));


            getEditingCell(childSheet.getRow(p.getRowIndex()),i++).setCellFormula(curMap.get("mostLoss"));


            int num = 0;
            int stopPos =0 ;//记录止损点

            for(int n= p.getPreTradePoint();n<(p.getRowIndex()+1);n++){

                getEditingCell(childSheet.getRow(n),18).setCellFormula(p.getTradeType()==TradeType.SHORT?"K"+(p.getRowIndex()+1)+"-D"+ (n+1) :"C"+ (n+1) +"-K" + (p.getRowIndex()+1));  //最多赚
                getEditingCell(childSheet.getRow(n),19).setCellFormula(p.getTradeType()==TradeType.LONG?"D"+ (n+1) +"-K" + (p.getRowIndex()+1):"K"+(p.getRowIndex()+1)+"-C"+ (n+1));  //最少赚
                getEditingCell(childSheet.getRow(n),20).setCellFormula(p.getTradeType()==TradeType.SHORT?"K"+(p.getRowIndex()+1)+"-E" + (n+1):"E"+ (n+1) +"-K"+(p.getRowIndex()+1));  //收盘赚


                boolean cond = p.getTradeType()==TradeType.SHORT && p.getBuyPrice() - getEditingCell(childSheet.getRow(n),2).getNumericCellValue() <= stopLine ;  //做空且亏尿
                boolean cond2 =p.getTradeType()==TradeType.LONG &&  getEditingCell(childSheet.getRow(n),3).getNumericCellValue() - p.getBuyPrice() <= stopLine ; //做多且亏尿

                if((cond || cond2) && num<1){  //找到亏尿点
                    stopPos = n+1;
                    num++;
                }

            }



            //填充亏尿点

            if(stopPos>0){ //存在亏尿点,填写亏尿前最多赚
                getEditingCell(childSheet.getRow(p.getRowIndex()),21).setCellFormula("max(S" +(p.getPreTradePoint()+1) + ":S" + stopPos + ")");
            }else{ //不存在，那就随便创建一个单元格,填充的是最多赚
                getEditingCell(childSheet.getRow(p.getRowIndex()),21).setCellFormula("max(S" + (p.getPreTradePoint()+1)  + ":S"+ (p.getRowIndex()+1) +")");

            }



            //以下是填充统计表

            Double curValue=0.0;
            Cell curCell=null;
            Row r= newSheet.createRow(newSheet.getLastRowNum()+1);


            FormulaEvaluator evaluator = null;
            try {
                evaluator = this.getFormulaEvalatorInstance(fileName,wbs);
            } catch (Exception e) {
                e.printStackTrace();
            }


            Cell sourceCell;
            for(Integer x = 0;x<22;x++){
                if(x==0) {
                    getEditingCell(r, x).setCellValue(
                            childSheet.getRow(p.getRowIndex()).getCell(x).getStringCellValue());
                }else{
                    curCell =getEditingCell(r, x);

                    sourceCell = childSheet.getRow(p.getRowIndex()).getCell(x);
                    if(sourceCell.getCellType() == Cell.CELL_TYPE_FORMULA){
                        curValue = getFormulaValue(childSheet.getRow(p.getRowIndex()).getCell(x),evaluator);

                    }else if(sourceCell.getCellType() == Cell.CELL_TYPE_NUMERIC){
                        curValue = sourceCell.getNumericCellValue();
                    }


                    if(curValue<0.0){

                        curCell.setCellStyle(style);
                    }

                    curCell.setCellValue(curValue);


                }
            }


        });


        //预计在此写统计总表


       //写入工作簿
        writeWbs(targetXlsFile,wbs);


    }

    private void writeWbs(File targetXlsFile,Workbook wbs) throws IOException{
        FileOutputStream out = new FileOutputStream(targetXlsFile);
        wbs.write(out);
        out.close();
    }

    private Cell getEditingCell(Row row,Integer i){
        Cell cell = row.getCell(i);
        if(cell==null){
            cell = row.createCell(i);
        }
        return cell;
    }


    private Workbook getWorkbookInstance(String fileName,InputStream is) throws Exception {
        if(fileName.matches(".+\\.(xls|XLS)$")) return new HSSFWorkbook(is);
        else if(fileName.matches(".+\\.(xlsx|XLSX)$")){
            return new XSSFWorkbook(is);
        }else{
            throw new IOException("只接受xls与xlsx文件");
        }
    }

    private FormulaEvaluator getFormulaEvalatorInstance(String fileName,Workbook wbs) throws Exception {
        if(fileName.matches(".+\\.(xls|XLS)$")) return new HSSFFormulaEvaluator((HSSFWorkbook) wbs);
        else if(fileName.matches(".+\\.(xlsx|XLSX)$")){
            return new XSSFFormulaEvaluator((XSSFWorkbook) wbs);
        }else{
            throw new IOException("只接受xls与xlsx文件");
        }
    }




    public static void main(String[] args) {
//        ExcelReader excelReader = new ExcelReader();
//        try {
//            File f =  new File("/Users/husu/Desktop/20140808.xls");
//            XlsData result = excelReader.loadXls(f,10);
//            List<MarketData> tradeList =excelReader.analyseData(result.getMdList(),result.getStopLossLine());
//            File newFile =excelReader.pasteFile(f,"");
//            excelReader.modify(newFile,tradeList);
//            System.out.println("========执行结束，文件地址" + newFile.getPath());
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
    }
}
