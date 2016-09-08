package com.parsec.sindle.model;

import java.util.List;

/**
 * @auther:husu
 * @version:1.0
 * @date 16/9/7.
 */
public class XlsData {
    private List<MarketData> mdList;
    private Double stopLossLine;

    public XlsData(List<MarketData> mdList, Double stopLossLine) {
        this.mdList = mdList;
        this.stopLossLine = stopLossLine;
    }

    public List<MarketData> getMdList() {
        return mdList;
    }

    public void setMdList(List<MarketData> mdList) {
        this.mdList = mdList;
    }

    public Double getStopLossLine() {
        return stopLossLine;
    }

    public void setStopLossLine(Double stopLossLine) {
        this.stopLossLine = stopLossLine;
    }
}
