using System;
using System.Collections.Generic;
using System.Linq;
using StockAnalyzer.Models;

namespace StockAnalyzer.Scoring
{
    public static class StockScorer
    {
        public static AnalysisResult Calculate(
            StockInfo info,
            FundamentalData fund,
            SectorFundamental secFund,
            List<InvestorDay> investors,   // 최신순 60일
            List<DailyBar> bars,        // 최신순 60일
            List<SectorSupplyDay> secSupply, // 최신순 20일
            ScoreConfig cfg)
        {
            var r = new AnalysisResult
            {
                Code = info.Code,
                Name = info.Name,
                Market = info.Market,
                SectorCode = info.SectorCode,
                SectorName = info.SectorName,
                CurrentPrice = info.CurrentPrice,
                Per = fund?.Per,
                Pbr = fund?.Pbr,
                Roe = fund?.Roe,
                SectorAvgPer = secFund?.AvgPer,
                SectorAvgPbr = secFund?.AvgPbr,
            };

            // ── 기업가치 점수 ────────────────────────────────────
            r.ValueScore = CalcValueScore(fund, secFund, cfg);

            // ── 종목수급 점수 ────────────────────────────────────
            r.StockSupplyScore = CalcStockSupplyScore(investors, bars, cfg, r);

            // ── 업종수급 점수 ────────────────────────────────────
            r.SectorSupplyScore = CalcSectorSupplyScore(secSupply, cfg, r);

            r.TotalScore = r.ValueScore + r.StockSupplyScore + r.SectorSupplyScore;
            r.Status = "완료";
            return r;
        }

        // ── 기업가치 ─────────────────────────────────────────────

        private static double CalcValueScore(
            FundamentalData fund, SectorFundamental sec, ScoreConfig cfg)
        {
            double score = 0;

            if (fund?.Per.HasValue == true && sec?.AvgPer.HasValue == true && sec.AvgPer > 0)
            {
                double discount = (sec.AvgPer.Value - fund.Per.Value) / sec.AvgPer.Value;
                score += Clamp(discount, 0, 1) * cfg.PerScore;
            }

            if (fund?.Pbr.HasValue == true && sec?.AvgPbr.HasValue == true && sec.AvgPbr > 0)
            {
                double discount = (sec.AvgPbr.Value - fund.Pbr.Value) / sec.AvgPbr.Value;
                score += Clamp(discount, 0, 1) * cfg.PbrScore;
            }

            if (fund?.Roe.HasValue == true && fund.Roe > 0)
                score += Clamp(fund.Roe.Value / 30.0, 0, 1) * cfg.RoeScore;

            return score;
        }

        // ── 종목수급 ─────────────────────────────────────────────

        private static double CalcStockSupplyScore(
            List<InvestorDay> investors, List<DailyBar> bars,
            ScoreConfig cfg, AnalysisResult r)
        {
            double score = 0;
            if (investors == null || investors.Count == 0) return 0;

            double maxForeign = investors.Max(d => Math.Abs((double)d.ForeignNet));
            double maxInst = investors.Max(d => Math.Abs((double)d.InstNet));
            if (maxForeign == 0) maxForeign = 1;
            if (maxInst == 0) maxInst = 1;

            // 당일
            r.ForeignNetD1 = investors.Count > 0 ? investors[0].ForeignNet : 0;
            r.InstNetD1 = investors.Count > 0 ? investors[0].InstNet : 0;

            // 5일 합
            r.ForeignNet5D = investors.Take(5).Sum(d => d.ForeignNet);
            r.InstNet5D = investors.Take(5).Sum(d => d.InstNet);

            // 10일 합
            r.ForeignNet10D = investors.Take(10).Sum(d => d.ForeignNet);
            r.InstNet10D = investors.Take(10).Sum(d => d.InstNet);

            // 20일 합
            r.ForeignNet20D = investors.Take(20).Sum(d => d.ForeignNet);
            r.InstNet20D = investors.Take(20).Sum(d => d.InstNet);

            score += NormalizeSupply(r.ForeignNetD1, maxForeign) * cfg.ForeignD1Score;
            score += NormalizeSupply(r.ForeignNet5D, maxForeign * 5) * cfg.Foreign5DScore;
            score += NormalizeSupply(r.ForeignNet20D, maxForeign * 20) * cfg.Foreign20DScore;
            score += NormalizeSupply(r.InstNetD1, maxInst) * cfg.InstD1Score;
            score += NormalizeSupply(r.InstNet5D, maxInst * 5) * cfg.Inst5DScore;
            score += NormalizeSupply(r.InstNet20D, maxInst * 20) * cfg.Inst20DScore;

            // 거래회전율 추세
            if (bars != null && bars.Count >= 20)
            {
                var bars20 = bars.Take(20).ToList();
                var bars60 = bars.Take(Math.Min(60, bars.Count)).ToList();

                double t20 = CalcAvgTurnover(bars20);
                double t60 = CalcAvgTurnover(bars60);
                r.Turnover20D = t20;
                r.Turnover60D = t60;

                if (t60 > 0)
                {
                    r.TurnoverRate = (t20 - t60) / t60 * 100.0;
                    double ratio = Clamp(r.TurnoverRate / cfg.TurnoverFullPct, 0, 1);
                    score += ratio * cfg.TurnoverScore;
                }
            }

            // 거래량 비교 (20D avg vs 60D avg)
            CalcVolumeTrend(bars, r);

            // 수급강도 추세
            CalcSupplyTrend(investors, bars, cfg, r);

            return score;
        }

        private static double CalcAvgTurnover(List<DailyBar> bars)
        {
            if (bars == null || bars.Count == 0) return 0;
            double sum = 0; int cnt = 0;
            foreach (var b in bars)
            {
                if (b.MarketCap > 0 && b.TradeAmount > 0)
                {
                    sum += b.TradeAmount / b.MarketCap;
                    cnt++;
                }
            }
            return cnt > 0 ? sum / cnt : 0;
        }

        /// <summary>60일 평균거래량 대비 20일 평균거래량 비교</summary>
        private static void CalcVolumeTrend(List<DailyBar> bars, AnalysisResult r)
        {
            if (bars == null || bars.Count < 20) return;

            var bars20 = bars.Take(20).ToList();
            var bars60 = bars.Take(Math.Min(60, bars.Count)).ToList();

            double avg20 = bars20.Average(b => (double)b.Volume);
            double avg60 = bars60.Average(b => (double)b.Volume);

            r.VolAvg20D = avg20;
            r.VolAvg60D = avg60;

            if (avg60 <= 0)
            {
                r.VolTrend = VolTrend.보합;
                return;
            }

            double ratio = (avg20 - avg60) / avg60;
            if (ratio > 0.10) r.VolTrend = VolTrend.상승;   // +10% 이상
            else if (ratio < -0.10) r.VolTrend = VolTrend.하락;   // -10% 이상
            else r.VolTrend = VolTrend.보합;
        }

        // 수급강도 추세 계산
        private static void CalcSupplyTrend(
            List<InvestorDay> investors, List<DailyBar> bars,
            ScoreConfig cfg, AnalysisResult r)
        {
            if (investors == null || investors.Count < 10) return;

            double s5 = CalcSupplyStrength(investors.Take(5).ToList(),
                                              bars?.Take(5).ToList());
            double sPrev = CalcSupplyStrength(investors.Skip(5).Take(5).ToList(),
                                              bars?.Skip(5).Take(5).ToList());

            r.SupplyStrength5D = s5;
            r.SupplyStrengthPrev5D = sPrev;

            double threshold = cfg.TrendThresholdPct / 100.0;

            if (sPrev <= 0 && s5 > 0)
                r.SupplyTrend = SupplyTrend.상승반전;
            else if (sPrev >= 0 && s5 < 0)
                r.SupplyTrend = SupplyTrend.하락반전;
            else
            {
                double change = Math.Abs(sPrev) > 0.001
                    ? (s5 - sPrev) / Math.Abs(sPrev)
                    : (s5 > 0 ? 1 : s5 < 0 ? -1 : 0);

                if (change > threshold) r.SupplyTrend = SupplyTrend.상승;
                else if (change < -threshold) r.SupplyTrend = SupplyTrend.하락;
                else r.SupplyTrend = SupplyTrend.보합;
            }
        }

        private static double CalcSupplyStrength(
            List<InvestorDay> days, List<DailyBar> bars)
        {
            if (days == null || days.Count == 0) return 0;

            long totalVol = days.Sum(d => d.Volume);
            long netQty = days.Sum(d => d.ForeignNet + d.InstNet);
            double qtyRatio = totalVol > 0 ? (double)netQty / totalVol : 0;

            double totalTrade = bars?.Sum(b => b.TradeAmount) ?? 0;
            double netAmt = days.Sum(d => (double)(d.ForeignNet + d.InstNet));
            double avgPrice = totalVol > 0 && totalTrade > 0
                ? totalTrade / totalVol : 0;
            double amtRatio = totalTrade > 0
                ? (netAmt * avgPrice) / totalTrade : 0;

            return (qtyRatio * 0.5) + (amtRatio * 0.5);
        }

        // ── 업종수급 (public) ────────────────────────────────────

        public static double CalcSectorSupplyPublic(
            List<SectorSupplyDay> supply, ScoreConfig cfg, AnalysisResult r)
            => CalcSectorSupplyScore(supply, cfg, r);

        // ── 업종수급 ─────────────────────────────────────────────

        private static double CalcSectorSupplyScore(
            List<SectorSupplyDay> supply, ScoreConfig cfg, AnalysisResult r)
        {
            if (supply == null || supply.Count == 0) return 0;

            double maxF = supply.Max(d => Math.Abs(d.ForeignNet));
            double maxI = supply.Max(d => Math.Abs(d.InstNet));
            if (maxF == 0) maxF = 1;
            if (maxI == 0) maxI = 1;

            r.SectorForeignD1 = supply.Count > 0 ? supply[0].ForeignNet : 0;
            r.SectorInstD1 = supply.Count > 0 ? supply[0].InstNet : 0;
            r.SectorForeign5D = supply.Take(5).Sum(d => d.ForeignNet);
            r.SectorInst5D = supply.Take(5).Sum(d => d.InstNet);
            r.SectorForeign20D = supply.Take(20).Sum(d => d.ForeignNet);
            r.SectorInst20D = supply.Take(20).Sum(d => d.InstNet);

            double score = 0;
            score += NormalizeSupply(r.SectorForeignD1, maxF) * cfg.SectorForeignD1Score;
            score += NormalizeSupply(r.SectorForeign5D, maxF * 5) * cfg.SectorForeign5DScore;
            score += NormalizeSupply(r.SectorForeign20D, maxF * 20) * cfg.SectorForeign20DScore;
            score += NormalizeSupply(r.SectorInstD1, maxI) * cfg.SectorInstD1Score;
            score += NormalizeSupply(r.SectorInst5D, maxI * 5) * cfg.SectorInst5DScore;
            score += NormalizeSupply(r.SectorInst20D, maxI * 20) * cfg.SectorInst20DScore;
            return score;
        }

        // ── 공통 유틸 ────────────────────────────────────────────

        private static double NormalizeSupply(double val, double maxAbs)
        {
            if (maxAbs == 0) return 0;
            return Clamp(val / maxAbs / 2.0 + 0.5, 0, 1);
        }

        private static double Clamp(double v, double min, double max)
            => v < min ? min : v > max ? max : v;
    }
}