using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AxKHOpenAPILib;
using StockAnalyzer.Api;
using StockAnalyzer.Models;
using StockAnalyzer.Scoring;
using StockAnalyzer.Utils;

namespace StockAnalyzer.Scoring
{
    public class AnalysisEngine
    {
        private readonly AxKHOpenAPI _ax;

        public event Action<string>           Log;
        public event Action<int, int, string> Progress;  // 현재, 전체, 종목명

        public AnalysisEngine(AxKHOpenAPI ax) => _ax = ax;

        public async Task<(
            List<AnalysisResult> Results,
            List<SectorSupplySummary> SectorKospi,
            List<SectorSupplySummary> SectorKosdaq)>
        RunAsync(List<string> codes, ScoreConfig cfg, CancellationToken ct)
        {
            var tradingDays60 = TradingDayHelper.GetRecentTradingDays(60);
            var tradingDays20 = tradingDays60.Take(20).ToList();
            var tradingDays5  = tradingDays60.Take(5).ToList();
            var fromDate60    = tradingDays60[tradingDays60.Count - 1];

            // KiwoomClient: 폼 수준 OCX를 공유하므로 한 번만 생성
            using (var kiwoom = new KiwoomClient(_ax))
            using (var krx    = new KrxClient(cfg.KrxAuthKey))
            {
                // ── 1. 종목 기본정보 (KRX Open API) ─────────────
                LogMsg("▶ 종목 기본정보 조회 (KRX Open API)...");
                var stockInfos = new Dictionary<string, StockInfo>();
                try
                {
                    var allStocks = await krx.GetAllStocksAsync();
                    LogMsg($"  KRX에서 {allStocks.Count}개 종목 로드 완료");

                    foreach (var code in codes)
                    {
                        ct.ThrowIfCancellationRequested();
                        var info = allStocks.FirstOrDefault(s => s.Code == code);
                        if (info != null) stockInfos[code] = info;
                    }
                    LogMsg($"  대상 종목 {stockInfos.Count}/{codes.Count}개 매칭");
                }
                catch (Exception ex)
                {
                    LogMsg($"  ✗ KRX 종목정보 조회 실패: {ex.Message}");
                    LogMsg("  → 키움 데이터로 대체합니다.");
                }

                // ── 2. 업종코드 수집 ─────────────────────────────
                var kospiSectors  = new HashSet<string>();
                var kosdaqSectors = new HashSet<string>();
                foreach (var info in stockInfos.Values)
                {
                    if (string.IsNullOrEmpty(info.SectorCode)) continue;
                    if (info.Market == "KOSDAQ") kosdaqSectors.Add(info.SectorCode);
                    else                          kospiSectors.Add(info.SectorCode);
                }

                // ── 3. 업종 PER/PBR (KRX) ───────────────────────
                LogMsg("▶ 업종 기본정보 조회...");
                var sectorFunds = new Dictionary<string, SectorFundamental>();
                try
                {
                    foreach (var sc in kospiSectors.Union(kosdaqSectors))
                    {
                        ct.ThrowIfCancellationRequested();
                        var market = kosdaqSectors.Contains(sc) ? "KOSDAQ" : "KOSPI";
                        var sf     = await krx.GetSectorFundamentalAsync(sc, market);
                        sectorFunds[sc] = sf;
                    }
                }
                catch (Exception ex)
                {
                    LogMsg($"  ✗ 업종정보 조회 실패: {ex.Message}");
                }

                // ── 4. 업종 수급 (키움 데이터로 보완) ────────────
                LogMsg("▶ 업종 수급 조회...");
                var sectorSupply = new Dictionary<string, List<SectorSupplyDay>>();
                foreach (var sc in kospiSectors.Union(kosdaqSectors))
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        var market = kosdaqSectors.Contains(sc) ? "KOSDAQ" : "KOSPI";
                        var data   = await krx.GetSectorSupplyAsync(sc, market, tradingDays20);
                        sectorSupply[sc] = data;
                    }
                    catch { sectorSupply[sc] = new List<SectorSupplyDay>(); }
                }

                // ── 5. 전체 업종 수급 현황 (3단 상단) ───────────
                LogMsg("▶ 전체 업종 수급 현황...");
                var sectorKospi  = new List<SectorSupplySummary>();
                var sectorKosdaq = new List<SectorSupplySummary>();
                try
                {
                    sectorKospi  = await krx.GetAllSectorSupplyAsync("KOSPI",  tradingDays5);
                    await Task.Delay(300, ct);
                    sectorKosdaq = await krx.GetAllSectorSupplyAsync("KOSDAQ", tradingDays5);
                }
                catch (Exception ex)
                {
                    LogMsg($"  ✗ 업종 현황 조회 실패: {ex.Message}");
                }

                // ── 6. 종목별 키움 데이터 + 점수화 ──────────────
                LogMsg($"▶ 종목별 분석 시작 ({codes.Count}개)...");
                var results = new List<AnalysisResult>();

                for (int i = 0; i < codes.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    var code = codes[i];

                    // KRX 정보가 없으면 키움에서 기본 정보 구성
                    if (!stockInfos.TryGetValue(code, out var info))
                    {
                        info = new StockInfo
                        {
                            Code = code, Name = code,
                            Market = "KOSPI", SectorCode = "", SectorName = ""
                        };
                        stockInfos[code] = info;
                    }

                    Progress?.Invoke(i + 1, codes.Count, info.Name);
                    LogMsg($"  [{i+1}/{codes.Count}] {info.Name} ({code})");

                    try
                    {
                        var fund      = await kiwoom.GetFundamentalAsync(code);
                        var investors = await kiwoom.GetInvestorDataAsync(code, fromDate60);
                        var bars      = await kiwoom.GetDailyBarAsync(code, fromDate60);

                        // 키움에서 이름을 업데이트
                        if (info.Name == code && fund != null)
                        {
                            // opt10001에서 종목명을 받아올 수 있으면 업데이트
                        }

                        sectorFunds.TryGetValue(info.SectorCode ?? "", out var sf);
                        sectorSupply.TryGetValue(info.SectorCode ?? "", out var ss);

                        var result = StockScorer.Calculate(
                            info, fund, sf, investors, bars,
                            ss ?? new List<SectorSupplyDay>(), cfg);
                        results.Add(result);
                        LogMsg($"    ✓ 총점={result.TotalScore:F1}");
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        results.Add(new AnalysisResult
                        { Code = code, Name = info.Name, Status = "오류", ErrorMsg = ex.Message });
                        LogMsg($"    ✗ {ex.Message}");
                    }
                }

                results.Sort((a, b) => b.TotalScore.CompareTo(a.TotalScore));
                LogMsg($"▶ 분석 완료: {results.Count}개 종목");
                return (results, sectorKospi, sectorKosdaq);
            }
        }

        private void LogMsg(string msg) => Log?.Invoke(msg);
    }
}
