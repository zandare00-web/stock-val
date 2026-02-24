using System;
using System.Collections.Generic;

namespace StockAnalyzer.Models
{
    public class StockCode
    {
        public string Code { get; set; }
    }

    public class StockInfo
    {
        public string Code         { get; set; }
        public string Name         { get; set; }
        public string Market       { get; set; }  // KOSPI / KOSDAQ
        public string SectorCode   { get; set; }  // KRX 업종코드
        public string SectorName   { get; set; }
        public double CurrentPrice { get; set; }
    }

    public class FundamentalData
    {
        public string  Code { get; set; }
        public double? Per  { get; set; }
        public double? Pbr  { get; set; }
        public double? Roe  { get; set; }
        public double? Eps  { get; set; }
        public double? Bps  { get; set; }
    }

    public class InvestorDay
    {
        public DateTime Date       { get; set; }
        public long     ForeignNet { get; set; }  // 외국인 순매수 수량
        public long     InstNet    { get; set; }  // 기관 순매수 수량
        public long     Volume     { get; set; }  // 거래량 (수급강도 계산용)
        public double   TradeAmount { get; set; } // 거래대금
    }

    public class DailyBar
    {
        public DateTime Date        { get; set; }
        public double   MarketCap   { get; set; }
        public double   TradeAmount { get; set; }
        public long     Volume      { get; set; }
    }

    public class SectorFundamental
    {
        public string  SectorCode { get; set; }
        public string  SectorName { get; set; }
        public double? AvgPer     { get; set; }
        public double? AvgPbr     { get; set; }
    }

    public class SectorSupplyDay
    {
        public DateTime Date       { get; set; }
        public double   ForeignNet { get; set; }  // 원
        public double   InstNet    { get; set; }  // 원
    }

    public enum SupplyTrend { 상승, 하락, 보합, 상승반전, 하락반전 }

    public class AnalysisResult
    {
        public string Code         { get; set; }
        public string Name         { get; set; }
        public string Market       { get; set; }
        public string SectorCode   { get; set; }
        public string SectorName   { get; set; }
        public double CurrentPrice { get; set; }
        public string Status       { get; set; }
        public string ErrorMsg     { get; set; }

        // 총점
        public double TotalScore        { get; set; }
        public double ValueScore        { get; set; }
        public double StockSupplyScore  { get; set; }
        public double SectorSupplyScore { get; set; }

        // 기업가치 세부
        public double? Per          { get; set; }
        public double? Pbr          { get; set; }
        public double? Roe          { get; set; }
        public double? SectorAvgPer { get; set; }
        public double? SectorAvgPbr { get; set; }

        // 종목수급 세부
        public long   ForeignNetD1  { get; set; }
        public long   ForeignNet5D  { get; set; }
        public long   ForeignNet20D { get; set; }
        public long   InstNetD1     { get; set; }
        public long   InstNet5D     { get; set; }
        public long   InstNet20D    { get; set; }

        // 거래회전율
        public double Turnover20D  { get; set; }
        public double Turnover60D  { get; set; }
        public double TurnoverRate { get; set; }  // 증가율(%)

        // 수급강도 추세
        public double      SupplyStrength5D     { get; set; }
        public double      SupplyStrengthPrev5D { get; set; }
        public SupplyTrend SupplyTrend          { get; set; }

        // 업종수급 세부
        public double SectorForeignD1  { get; set; }
        public double SectorForeign5D  { get; set; }
        public double SectorForeign20D { get; set; }
        public double SectorInstD1     { get; set; }
        public double SectorInst5D     { get; set; }
        public double SectorInst20D    { get; set; }
    }

    public class SectorSupplySummary
    {
        public string SectorCode   { get; set; }
        public string SectorName   { get; set; }
        public string Market       { get; set; }
        public double ForeignNet5D { get; set; }
        public double InstNet5D    { get; set; }
        public double TotalNet5D   { get; set; }  // 합산 (정렬용)
    }
}
