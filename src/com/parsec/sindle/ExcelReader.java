package com.parsec.sindle;

import com.parsec.sindle.model.MarketData;
import com.parsec.sindle.model.TradeType;
import com.parsec.sindle.model.XlsData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
    /**
     * 从文件中读取行情数据<br>
     * @param f
     * @return
     * @throws IOException
     */
    public XlsData loadXls(File f) throws Exception {

        InputStream is = new FileInputStream(f);
        org.apache.poi.ss.usermodel.Workbook wbs;

        String fileName = f.getName();

        wbs = this.getWorkbookInstance(fileName,is);

        org.apache.poi.ss.usermodel.Sheet childSheet = wbs.getSheetAt(0);


        Double mostLossLine = java.lang.Math.abs(childSheet.getRow(2).getCell(1).getNumericCellValue())*(-1);

        MarketData marketData;
        Double preDk = null;
        Row curRow;


        List<MarketData> mdList= new ArrayList<>();


        for(int i=childSheet.getLastRowNum();i>=5;i--){
            curRow = childSheet.getRow(i);

            boolean flag = false;

            if(preDk==null){
                preDk = curRow.getCell(9).getNumericCellValue();//获得多空状态
            }else if(preDk != curRow.getCell(9).getNumericCellValue()){
                preDk = curRow.getCell(9).getNumericCellValue();
                flag =true;
            }

            if(i==childSheet.getLastRowNum()){
                preDk = curRow.getCell(9).getNumericCellValue();
                flag=true;
            }

            marketData = new MarketData(i);


            marketData.setTradePoint(flag);
            marketData.setTradeType(preDk==0?TradeType.SHORT:TradeType.LONG);

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


    public List<MarketData> analyseData(List<MarketData> mdList,Double mostLossLine){
        List<MarketData> tradeList = mdList.stream().sorted((p1, p2)->(p1.getRowIndex()>p2.getRowIndex()?1:-1)).filter(MarketData::getTradePoint).collect(Collectors.toList());

        tradeList.stream().reduce(new MarketData(4),(p1,p2)-> {  //这个地方的初始值是4，是因为下面做了加1操作
            p2.setPreTradePoint(p1.getRowIndex()+1);
            return p2;
        });

        tradeList.forEach(p->{
            Map<String,Double> map =  new HashMap<>();

            //开仓点位
            map.put("openPoint",mdList.stream().filter(md-> {
                int pp = p.getPreTradePoint()-1;
                pp = pp<5?5:pp;
                return md.getRowIndex() == pp;
            }).max(Comparator.comparing(MarketData::getClosePrice)).get().getClosePrice());

            //平仓点位
            map.put("sellPoint",p.getClosePrice());

            //最高价
            map.put("highestPrice",mdList.stream().filter(md-> (md.getRowIndex()>=(p.getPreTradePoint()) && md.getRowIndex()<=p.getRowIndex()))
                    .max(Comparator.comparing(MarketData::getHightestPrice)).get().getHightestPrice());

            //最低价
            map.put("lowestPrice",mdList.stream().filter(md-> (md.getRowIndex()>=(p.getPreTradePoint()) && md.getRowIndex()<=p.getRowIndex()))
                    .min(Comparator.comparing(MarketData::getLowestPrice)).get().getLowestPrice());

            //结果无止损
            Double lossNoStop = (p.getTradeType()==TradeType.SHORT)?map.get("openPoint")-map.get("sellPoint"):map.get("sellPoint")-map.get("openPoint");
            map.put("lossNoStop",lossNoStop);

            //计算最大亏损
            map.put("mostLoss",p.getTradeType()==TradeType.LONG?(map.get("lowestPrice")-map.get("openPoint")):(map.get("openPoint")-map.get("highestPrice")));

            //结果有止损   逻辑是，如果没有止损，则为无止损结果，止损，则为止损值
            map.put("lossStop",map.get("mostLoss")<=mostLossLine?mostLossLine:lossNoStop);

            //计算最大赢利
            map.put("mostEarn",p.getTradeType()==TradeType.LONG?(map.get("highestPrice")-map.get("openPoint")):(map.get("openPoint")-map.get("lowestPrice")));

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
     * @param sourceFile
     * @return
     * @throws IOException
     */
    public File pasteFile(File sourceFile) throws IOException {

        String sourceFileName = sourceFile.getName();
        String targetFilePath = sourceFile.getParentFile().getPath() + File.separator + sourceFileName.substring(0,sourceFileName.lastIndexOf("."))
              + "-" + (new Date()).getTime() + "(done)"  + "." + sourceFileName.substring(sourceFileName.lastIndexOf(".")+1);


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
    public void modify(File targetXlsFile,List<MarketData> tradeList) throws IOException {

        InputStream is = new FileInputStream(targetXlsFile);


        String fileName = targetXlsFile.getName();

        Workbook wbs = this.getWorkbookInstance(fileName,is);

        Sheet childSheet = wbs.getSheetAt(0);

        tradeList.forEach(p->{          //填充交易点
            Map<String,Double> curMap = p.getResultMap();
            int i= 10;
            childSheet.getRow(p.getRowIndex()).getCell(i++).setCellValue(curMap.get("openPoint"));
            childSheet.getRow(p.getRowIndex()).getCell(i++).setCellValue(curMap.get("sellPoint"));
            childSheet.getRow(p.getRowIndex()).getCell(i++).setCellValue(curMap.get("lossNoStop"));
            childSheet.getRow(p.getRowIndex()).getCell(i++).setCellValue(curMap.get("lossStop"));//结果有止损
            childSheet.getRow(p.getRowIndex()).getCell(i++).setCellValue(curMap.get("highestPrice")); //最高价
            childSheet.getRow(p.getRowIndex()).getCell(i++).setCellValue(curMap.get("lowestPrice"));
            childSheet.getRow(p.getRowIndex()).getCell(i++).setCellValue(curMap.get("mostEarn"));
            childSheet.getRow(p.getRowIndex()).getCell(i++).setCellValue(curMap.get("mostLoss"));
            int dk = 1; //这个参数就是这么屌
            if(p.getTradeType()==TradeType.SHORT){
                dk = -1;
            }

            Double curHOL ;
            Double curHOL4zsz;//最高价或者最低价，对应最少赚
            Double closePrice;//收盘价，对应收盘赚
            for(int n= p.getPreTradePoint();n<(p.getRowIndex()+1);n++){
                curHOL = (p.getTradeType()==TradeType.LONG) ? childSheet.getRow(n).getCell(2).getNumericCellValue():childSheet.getRow(n).getCell(3).getNumericCellValue();    //多取最高价计算  空取最低价计算  用来计算最多赚
                curHOL4zsz = (p.getTradeType()==TradeType.SHORT) ? childSheet.getRow(n).getCell(2).getNumericCellValue():childSheet.getRow(n).getCell(3).getNumericCellValue();    //多取最高价计算  空取最低价计算  用来计算最多赚
                closePrice = childSheet.getRow(n).getCell(4).getNumericCellValue();
                childSheet.getRow(n).getCell(18).setCellValue((curHOL-p.getResultMap().get("openPoint"))*dk);  //最多赚，屌不屌
                childSheet.getRow(n).getCell(19).setCellValue((curHOL4zsz-p.getResultMap().get("openPoint"))*dk);  //最少赚，屌不屌
                childSheet.getRow(n).getCell(20).setCellValue((closePrice-p.getResultMap().get("openPoint"))*dk);  //最少赚，屌不屌

            }

        });

        FileOutputStream out = null;
        try {
            out = new FileOutputStream(targetXlsFile);
            wbs.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if(out!=null){
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }


    private Workbook getWorkbookInstance(String fileName,InputStream is) throws IOException {
        if(fileName.matches(".+\\.(xls|XLS)$")) return new HSSFWorkbook(is);
        else if(fileName.matches(".+\\.(xlsx|XLSX)$")){
            return new XSSFWorkbook(is);
        }else{
            throw new IOException("只接受xls与xlsx文件");
        }
    }

    public static void main(String[] args) {
        ExcelReader excelReader = new ExcelReader();
        try {
            File f =  new File("/Users/husu/Desktop/20140808.xls");
            XlsData result = excelReader.loadXls(f);
            List<MarketData> tradeList =excelReader.analyseData(result.getMdList(),result.getStopLossLine());
            File newFile =excelReader.pasteFile(f);
            excelReader.modify(newFile,tradeList);
            System.out.println("========执行结束，文件地址" + newFile.getPath());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
