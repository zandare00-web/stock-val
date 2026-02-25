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
            List<InvestorDay> investors,
            List<DailyBar> bars,
            List<SectorSupplyDay> secSupply,
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

            r.ValueScore = CalcValueScore(fund, secFund, cfg);
            r.StockSupplyScore = CalcStockSupplyScore(investors, bars, cfg, r);
            r.SectorSupplyScore = CalcSectorSupplyScore(secSupply, cfg, r);

            r.TotalScore = r.ValueScore + r.StockSupplyScore + r.SectorSupplyScore;
            r.Status = "완료";
            return r;
        }

        // ── 기업가치 (기본 5점) ──────────────────────────────────

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

        // ── 종목수급 (기본 55점) ─────────────────────────────────

        private static double CalcStockSupplyScore(
            List<InvestorDay> investors, List<DailyBar> bars,
            ScoreConfig cfg, AnalysisResult r)
        {
            double score = 0;
            if (investors == null || investors.Count == 0) return 0;

            // 금액 기록 (점수계산 + UI 표시용)
            r.ForeignNetAmtD1 = investors.Count > 0 ? investors[0].ForeignNetAmt : 0;
            r.InstNetAmtD1 = investors.Count > 0 ? investors[0].InstNetAmt : 0;
            r.ForeignNetAmt5D = investors.Take(5).Sum(d => d.ForeignNetAmt);
            r.InstNetAmt5D = investors.Take(5).Sum(d => d.InstNetAmt);
            r.ForeignNetAmt10D = investors.Take(10).Sum(d => d.ForeignNetAmt);
            r.InstNetAmt10D = investors.Take(10).Sum(d => d.InstNetAmt);
            r.ForeignNetAmt20D = investors.Take(20).Sum(d => d.ForeignNetAmt);
            r.InstNetAmt20D = investors.Take(20).Sum(d => d.InstNetAmt);

            // 수량 기록 (UI 표시용)
            r.ForeignNetQtyD1 = investors.Count > 0 ? investors[0].ForeignNetQty : 0;
            r.InstNetQtyD1 = investors.Count > 0 ? investors[0].InstNetQty : 0;
            r.ForeignNetQty5D = investors.Take(5).Sum(d => d.ForeignNetQty);
            r.InstNetQty5D = investors.Take(5).Sum(d => d.InstNetQty);
            r.ForeignNetQty10D = investors.Take(10).Sum(d => d.ForeignNetQty);
            r.InstNetQty10D = investors.Take(10).Sum(d => d.InstNetQty);
            r.ForeignNetQty20D = investors.Take(20).Sum(d => d.ForeignNetQty);
            r.InstNetQty20D = investors.Take(20).Sum(d => d.InstNetQty);

            // ★ 금액 기준 외국인+기관 합계로 점수 계산
            long combined5D = r.ForeignNetAmt5D + r.InstNetAmt5D;
            long combined10D = r.ForeignNetAmt10D + r.InstNetAmt10D;
            long combined20D = r.ForeignNetAmt20D + r.InstNetAmt20D;

            // 정규화 기준: 각 기간의 일별 합계(외+기) 최대 절대값
            var dailyCombined = investors.Select(d => (double)(d.ForeignNetAmt + d.InstNetAmt)).ToList();
            double maxDaily = dailyCombined.Count > 0
                ? dailyCombined.Max(v => Math.Abs(v)) : 1;
            if (maxDaily == 0) maxDaily = 1;

            score += NormalizeSupply(combined5D, maxDaily * 5) * cfg.Supply5DScore;
            score += NormalizeSupply(combined10D, maxDaily * 10) * cfg.Supply10DScore;
            score += NormalizeSupply(combined20D, maxDaily * 20) * cfg.Supply20DScore;

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

            // 거래량 비교
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
                { sum += b.TradeAmount / b.MarketCap; cnt++; }
            }
            return cnt > 0 ? sum / cnt : 0;
        }

        private static void CalcVolumeTrend(List<DailyBar> bars, AnalysisResult r)
        {
            if (bars == null || bars.Count < 20) return;
            double avg20 = bars.Take(20).Average(b => (double)b.Volume);
            double avg60 = bars.Take(Math.Min(60, bars.Count)).Average(b => (double)b.Volume);
            r.VolAvg20D = avg20;
            r.VolAvg60D = avg60;
            if (avg60 <= 0) { r.VolTrend = VolTrend.보합; return; }
            double ratio = (avg20 - avg60) / avg60;
            r.VolTrend = ratio > 0.10 ? VolTrend.상승 : ratio < -0.10 ? VolTrend.하락 : VolTrend.보합;
        }

        private static void CalcSupplyTrend(
            List<InvestorDay> investors, List<DailyBar> bars,
            ScoreConfig cfg, AnalysisResult r)
        {
            if (investors == null || investors.Count < 10) return;
            double s5 = CalcSupplyStrength(investors.Take(5).ToList(), bars?.Take(5).ToList());
            double sPrev = CalcSupplyStrength(investors.Skip(5).Take(5).ToList(), bars?.Skip(5).Take(5).ToList());
            r.SupplyStrength5D = s5;
            r.SupplyStrengthPrev5D = sPrev;

            double threshold = cfg.TrendThresholdPct / 100.0;
            if (sPrev <= 0 && s5 > 0) r.SupplyTrend = SupplyTrend.상승반전;
            else if (sPrev >= 0 && s5 < 0) r.SupplyTrend = SupplyTrend.하락반전;
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

        private static double CalcSupplyStrength(List<InvestorDay> days, List<DailyBar> bars)
        {
            if (days == null || days.Count == 0) return 0;

            // 수급강도 = (외인+기관 순매수금액 합계) / (동기간 거래대금 합계)
            // ForeignNetAmt/InstNetAmt: 원, DailyBar.TradeAmount: 원
            double totalTrade = bars?.Sum(b => b.TradeAmount) ?? 0;
            if (totalTrade <= 0) return 0;

            double netAmt = days.Sum(d => (double)(d.ForeignNetAmt + d.InstNetAmt));
            return netAmt / totalTrade;
        }

        // ── 업종수급 (기본 40점) ─────────────────────────────────

        public static double CalcSectorSupplyPublic(
            List<SectorSupplyDay> supply, ScoreConfig cfg, AnalysisResult r)
            => CalcSectorSupplyScore(supply, cfg, r);

        private static double CalcSectorSupplyScore(
            List<SectorSupplyDay> supply, ScoreConfig cfg, AnalysisResult r)
        {
            if (supply == null || supply.Count == 0) return 0;

            // 개별 값 기록 (UI 표시용)
            r.SectorForeignD1 = supply.Count > 0 ? supply[0].ForeignNet : 0;
            r.SectorInstD1 = supply.Count > 0 ? supply[0].InstNet : 0;
            r.SectorForeign5D = supply.Take(5).Sum(d => d.ForeignNet);
            r.SectorInst5D = supply.Take(5).Sum(d => d.InstNet);
            r.SectorForeign20D = supply.Take(20).Sum(d => d.ForeignNet);
            r.SectorInst20D = supply.Take(20).Sum(d => d.InstNet);

            // ★ 외국인+기관 합계로 점수 계산
            double combined5D = supply.Take(5).Sum(d => d.ForeignNet + d.InstNet);
            double combined10D = supply.Take(10).Sum(d => d.ForeignNet + d.InstNet);

            var dailyCombined = supply.Select(d => Math.Abs(d.ForeignNet + d.InstNet)).ToList();
            double maxDaily = dailyCombined.Count > 0 ? dailyCombined.Max() : 1;
            if (maxDaily == 0) maxDaily = 1;

            double score = 0;
            score += NormalizeSupply(combined5D, maxDaily * 5) * cfg.SectorSupply5DScore;
            score += NormalizeSupply(combined10D, maxDaily * 10) * cfg.SectorSupply10DScore;
            return score;
        }

        // ── 유틸 ────────────────────────────────────────────────

        private static double NormalizeSupply(double val, double maxAbs)
        {
            if (maxAbs == 0) return 0;
            return Clamp(val / maxAbs / 2.0 + 0.5, 0, 1);
        }

        private static double Clamp(double v, double min, double max)
            => v < min ? min : v > max ? max : v;
    }
}