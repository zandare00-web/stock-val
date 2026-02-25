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
                // 음수 PER(적자 기업)은 저평가가 아니므로 0점 처리
                if (fund.Per.Value > 0)
                {
                    double discount = (sec.AvgPer.Value - fund.Per.Value) / sec.AvgPer.Value;
                    score += Clamp(discount, 0, 1) * cfg.PerScore;
                }
            }

            if (fund?.Pbr.HasValue == true && sec?.AvgPbr.HasValue == true && sec.AvgPbr > 0)
            {
                // 음수 PBR(자본잠식 기업)은 저평가가 아니므로 0점 처리
                if (fund.Pbr.Value > 0)
                {
                    double discount = (sec.AvgPbr.Value - fund.Pbr.Value) / sec.AvgPbr.Value;
                    score += Clamp(discount, 0, 1) * cfg.PbrScore;
                }
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

            // 개별 값 기록 (UI 표시용) - 수량 기준(주)
            r.ForeignNetD1 = investors.Count > 0 ? investors[0].ForeignNetQty : 0;
            r.InstNetD1 = investors.Count > 0 ? investors[0].InstNetQty : 0;
            r.ForeignNet5D = investors.Take(5).Sum(d => d.ForeignNetQty);
            r.InstNet5D = investors.Take(5).Sum(d => d.InstNetQty);
            r.ForeignNet10D = investors.Take(10).Sum(d => d.ForeignNetQty);
            r.InstNet10D = investors.Take(10).Sum(d => d.InstNetQty);
            r.ForeignNet20D = investors.Take(20).Sum(d => d.ForeignNetQty);
            r.InstNet20D = investors.Take(20).Sum(d => d.InstNetQty);

            // 종목별 수급현황은 금액을 사용하지 않고 수량(주)만 사용
            long combined5D = r.ForeignNet5D + r.InstNet5D;
            long combined10D = r.ForeignNet10D + r.InstNet10D;
            long combined20D = r.ForeignNet20D + r.InstNet20D;

            // 정규화 기준: 각 기간의 일별 합계(외+기) 최대 절대값 (수량)
            var dailyCombined = investors.Select(d => (double)(d.ForeignNetQty + d.InstNetQty)).ToList();
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
            double s5 = CalcSupplyStrength(investors.Take(5).ToList(), bars);
            double sPrev = CalcSupplyStrength(investors.Skip(5).Take(5).ToList(), bars);
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

            // 수급강도(수량 기준) = (외인+기관 순매수수량 합계) / (동기간 거래량 합계)
            double totalVolume = 0;

            // opt10059 응답의 누적거래량이 있으면 우선 사용
            if (days.Any(d => d.Volume > 0))
                totalVolume = days.Sum(d => Math.Abs((double)d.Volume));

            // 없으면 일봉 거래량으로 보완
            if (totalVolume <= 0 && bars != null && bars.Count > 0)
            {
                var daySet = new HashSet<DateTime>(days.Select(x => x.Date.Date));
                totalVolume = bars.Where(b => daySet.Contains(b.Date.Date)).Sum(b => Math.Abs((double)b.Volume));
            }

            if (totalVolume <= 0) return 0;

            double netQty = days.Sum(d => (double)(d.ForeignNetQty + d.InstNetQty));
            return netQty / totalVolume;
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