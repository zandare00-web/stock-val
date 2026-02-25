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
        public string Code { get; set; }
        public string Name { get; set; }
        public string Market { get; set; }  // KOSPI / KOSDAQ
        public string SectorCode { get; set; }
        public string SectorName { get; set; }
        public double CurrentPrice { get; set; }
    }

    public class FundamentalData
    {
        public string Code { get; set; }
        public double? Per { get; set; }
        public double? Pbr { get; set; }
        public double? Roe { get; set; }
        public double? Eps { get; set; }
        public double? Bps { get; set; }
        public double MarketCap { get; set; }  // 시가총액 (원)
    }

    public class InvestorDay
    {
        public DateTime Date { get; set; }

        // 키움 opt10059 실제 수신값 (정확값 분리 보관)
        public long ForeignNetQty { get; set; }   // 외국인 순매수 수량(주)
        public long InstNetQty { get; set; }      // 기관 순매수 수량(주)
        public long ForeignNetAmt { get; set; }   // 외국인 순매수 금액(원)
        public long InstNetAmt { get; set; }      // 기관 순매수 금액(원)

        // 하위 호환(기존 로직): 기본은 금액(원)
        public long ForeignNet { get; set; }
        public long InstNet { get; set; }

        public long Volume { get; set; }  // 누적거래량 (주)
        public double TradeAmount { get; set; } // 거래대금 (원)
    }

    public class DailyBar
    {
        public DateTime Date { get; set; }
        public double MarketCap { get; set; }
        public double TradeAmount { get; set; }
        public long Volume { get; set; }
    }

    public class SectorFundamental
    {
        public string SectorCode { get; set; }
        public string SectorName { get; set; }
        public double? AvgPer { get; set; }
        public double? AvgPbr { get; set; }
    }

    public class SectorSupplyDay
    {
        public DateTime Date { get; set; }
        public double ForeignNet { get; set; }  // 원
        public double InstNet { get; set; }  // 원
    }

    /// <summary>opt10051 업종별투자자순매수 1행</summary>
    public class SectorInvestorRow
    {
        public string SectorCode { get; set; }
        public string SectorName { get; set; }
        public double ForeignNet { get; set; }
        public double InstNet { get; set; }
    }

    public enum SupplyTrend { 상승, 하락, 보합, 상승반전, 하락반전 }
    public enum VolTrend { 상승, 하락, 보합 }

    public class AnalysisResult
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string Market { get; set; }
        public string SectorCode { get; set; }
        public string SectorName { get; set; }
        public double CurrentPrice { get; set; }
        public string Status { get; set; }
        public string ErrorMsg { get; set; }

        // 총점
        public double TotalScore { get; set; }
        public double ValueScore { get; set; }
        public double StockSupplyScore { get; set; }
        public double SectorSupplyScore { get; set; }

        // 기업가치 세부
        public double? Per { get; set; }
        public double? Pbr { get; set; }
        public double? Roe { get; set; }
        public double? SectorAvgPer { get; set; }
        public double? SectorAvgPbr { get; set; }

        // 종목수급 세부 (수량 기준, 주)
        public long ForeignNetD1 { get; set; }
        public long ForeignNet5D { get; set; }
        public long ForeignNet10D { get; set; }
        public long ForeignNet20D { get; set; }
        public long InstNetD1 { get; set; }
        public long InstNet5D { get; set; }
        public long InstNet10D { get; set; }
        public long InstNet20D { get; set; }

        // 종목수급 세부 (금액 기준, 원)
        public long ForeignNetAmtD1 { get; set; }
        public long ForeignNetAmt5D { get; set; }
        public long ForeignNetAmt10D { get; set; }
        public long ForeignNetAmt20D { get; set; }
        public long InstNetAmtD1 { get; set; }
        public long InstNetAmt5D { get; set; }
        public long InstNetAmt10D { get; set; }
        public long InstNetAmt20D { get; set; }

        // 거래회전율
        public double Turnover20D { get; set; }
        public double Turnover60D { get; set; }
        public double TurnoverRate { get; set; }

        // 거래량 비교 (20D avg vs 60D avg)
        public double VolAvg20D { get; set; }
        public double VolAvg60D { get; set; }
        public VolTrend VolTrend { get; set; }

        // 수급강도 추세
        public double SupplyStrength5D { get; set; }
        public double SupplyStrengthPrev5D { get; set; }
        public SupplyTrend SupplyTrend { get; set; }

        // 업종수급 세부
        public double SectorForeignD1 { get; set; }
        public double SectorForeign5D { get; set; }
        public double SectorForeign20D { get; set; }
        public double SectorInstD1 { get; set; }
        public double SectorInst5D { get; set; }
        public double SectorInst20D { get; set; }
    }

    public class SectorSupplySummary
    {
        public string SectorCode { get; set; }
        public string SectorName { get; set; }
        public string Market { get; set; }
        public double ForeignNet5D { get; set; }
        public double InstNet5D { get; set; }
        public double TotalNet5D { get; set; }
    }

    /// <summary>컨센서스 데이터</summary>
    public class ConsensusData
    {
        public string Code { get; set; }
        public string Opinion { get; set; }
        public double? TargetPrice { get; set; }
        public double? TargetPriceMin { get; set; }
        public double? TargetPriceMax { get; set; }
        public double? CurrentPrice { get; set; }
        public double? DeviationPct { get; set; }
        public double? ConsensusPer { get; set; }
        public double? ConsensusEps { get; set; }
        public int AnalystCount { get; set; }
        public string LatestReportDate { get; set; }
        public string Source { get; set; }
        public DateTime FetchedAt { get; set; }
        public bool IsValid => !string.IsNullOrEmpty(Opinion) || TargetPrice.HasValue;
    }
}