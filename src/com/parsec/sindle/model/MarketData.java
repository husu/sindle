package com.parsec.sindle.model;

import com.sun.org.apache.xpath.internal.operations.Bool;

import java.util.Map;

/**
 * @auther:husu
 * @version:1.0
 * @date 16/9/6.
 */
public class MarketData {
    private Double openPrice ;
    private Double hightestPrice;
    private Double lowestPrice;
    private Double closePrice;
    private int rowIndex = 0;
    private Boolean isTradePoint; //是否产生交易的时间点（行）
    private TradeType tradeType;
    private Integer preTradePoint;
    private Double buyPrice;//开仓点位
    private Map<String,String> resultMap;
    private Boolean stopped;  //有止损

    public Map<String, String> getResultMap() {
        return resultMap;
    }


    public Boolean getStopped() {
        return stopped;
    }

    public void setStopped(Boolean stopped) {
        this.stopped = stopped;
    }

    public Double getBuyPrice() {
        return buyPrice;
    }

    public void setBuyPrice(Double buyPrice) {
        this.buyPrice = buyPrice;
    }

    public void setResultMap(Map<String, String> resultMap) {
        this.resultMap = resultMap;
    }

    public Integer getPreTradePoint() {
        return preTradePoint;
    }

    public void setPreTradePoint(Integer preTradePoint) {
        this.preTradePoint = preTradePoint;
    }

    public Boolean getTradePoint() {
        return isTradePoint;
    }

    public void setTradePoint(Boolean tradePoint) {
        isTradePoint = tradePoint;
    }

    public TradeType getTradeType() {
        return tradeType;
    }

    public void setTradeType(TradeType tradeType) {
        this.tradeType = tradeType;
    }

    public MarketData(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public Double getOpenPrice() {
        return openPrice;
    }

    public void setOpenPrice(Double openPrice) {
        this.openPrice = openPrice;
    }

    public Double getHightestPrice() {
        return hightestPrice;
    }

    public void setHightestPrice(Double hightestPrice) {
        this.hightestPrice = hightestPrice;
    }

    public Double getLowestPrice() {
        return lowestPrice;
    }

    public void setLowestPrice(Double lowestPrice) {
        this.lowestPrice = lowestPrice;
    }

    public Double getClosePrice() {
        return closePrice;
    }

    public void setClosePrice(Double closePrice) {
        this.closePrice = closePrice;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }
}
