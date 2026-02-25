using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AxKHOpenAPILib;
using StockAnalyzer.Api;
using StockAnalyzer.Models;
using StockAnalyzer.Utils;

namespace StockAnalyzer.Scoring
{
    public class AnalysisEngine
    {
        private readonly AxKHOpenAPI _ax;

        public event Action<string> Log;
        public event Action<int, int, string> Progress;

        public AnalysisEngine(AxKHOpenAPI ax) => _ax = ax;

        public async Task<(
            List<AnalysisResult> Results,
            List<SectorSupplySummary> SectorKospi,
            List<SectorSupplySummary> SectorKosdaq)>
        RunAsync(List<string> codes, ScoreConfig cfg, CancellationToken ct)
        {
            var tradingDays60 = TradingDayHelper.GetRecentTradingDays(60);
            var fromDate60 = tradingDays60[tradingDays60.Count - 1];

            using (var kiwoom = new KiwoomClient(_ax))
            using (var krx = new KrxClient(cfg.KrxAuthKey))
            {
                // ── 1. KRX 종목정보 (이름/시장 기본) ─────────────
                LogMsg("▶ 종목 기본정보 조회 (KRX Open API)...");
                var stockInfos = new Dictionary<string, StockInfo>();
                try
                {
                    var allStocks = await krx.GetAllStocksAsync();
                    LogMsg($"  KRX에서 {allStocks.Count}개 종목 로드 완료");

                    var allStocksByCode = allStocks
                        .GroupBy(s => s.Code)
                        .ToDictionary(g => g.Key, g => g.First());

                    foreach (var code in codes)
                    {
                        ct.ThrowIfCancellationRequested();
                        if (allStocksByCode.TryGetValue(code, out var info) && info != null)
                            stockInfos[code] = info;
                    }
                    LogMsg($"  대상 종목 {stockInfos.Count}/{codes.Count}개 매칭");
                }
                catch (Exception ex)
                {
                    LogMsg($"  ✗ KRX 조회 실패: {ex.Message}");
                }

                // ── 2. 키움 KOA_Functions → 업종이름/시장 ────────
                LogMsg("▶ 종목별 업종 정보 수집 (키움 KOA_Functions)...");
                var stockSectorMap = new Dictionary<string, string>();  // code → 업종이름
                var stockMarketMap = new Dictionary<string, string>();  // code → KOSPI/KOSDAQ
                foreach (var code in codes)
                {
                    try
                    {
                        string raw = _ax.KOA_Functions("GetMasterStockInfo", code);
                        var parsed = ParseStockInfo(raw);
                        if (parsed.SectorName != null) stockSectorMap[code] = parsed.SectorName;
                        if (parsed.Market != null) stockMarketMap[code] = parsed.Market;
                    }
                    catch { }
                }
                LogMsg($"  업종정보: {stockSectorMap.Count}/{codes.Count}개 수집");

                // StockInfo에 업종 반영
                foreach (var code in codes)
                {
                    if (!stockInfos.TryGetValue(code, out var info))
                    {
                        info = new StockInfo
                        {
                            Code = code,
                            Name = _ax.GetMasterCodeName(code) ?? code,
                            Market = "KOSPI",
                            SectorCode = "",
                            SectorName = ""
                        };
                        stockInfos[code] = info;
                    }
                    if (stockSectorMap.TryGetValue(code, out var secName))
                        info.SectorName = secName;
                    if (stockMarketMap.TryGetValue(code, out var mkt))
                        info.Market = mkt;
                }

                // ── 3. 업종수급 조회 (opt10051) ──────────────────
                // 중요: 업종코드는 시장 간 중복될 수 있어 (시장,업종코드) 복합키 사용
                LogMsg("▶ 업종별 투자자순매수 조회 (opt10051)...");
                var sectorSupply = new Dictionary<string, List<SectorSupplyDay>>();      // key=KOSPI|027
                var sectorNameMap = new Dictionary<string, string>();                     // key → name
                var sectorMarketMap = new Dictionary<string, string>();                    // key → KOSPI/KOSDAQ
                var actualTradingDays = new List<DateTime>();
                try
                {
                    const int TARGET_DAYS = 20;
                    const int MAX_TRIES = 45;
                    var d = TradingDayHelper.GetLatestCompletedTradingDay();

                    int tries = 0;
                    int emptyCount = 0;
                    while (actualTradingDays.Count < TARGET_DAYS && tries < MAX_TRIES)
                    {
                        ct.ThrowIfCancellationRequested();
                        tries++;

                        var dTrade = TradingDayHelper.GetPreviousTradingDay(d);
                        var dateStr = dTrade.ToString("yyyyMMdd");

                        var kospiRows = await kiwoom.GetSectorInvestorAsync("0", dateStr, "0");
                        if (kospiRows.Count == 0)
                        {
                            emptyCount++;
                            if (emptyCount <= 5)
                                LogMsg($"  ⚠ {dateStr}: 코스피 0건 (휴장/데이터없음 추정, 건너뜀 {emptyCount}회)");

                            d = dTrade.AddDays(-1);
                            continue;
                        }

                        actualTradingDays.Add(dTrade);
                        var kosdaqRows = await kiwoom.GetSectorInvestorAsync("1", dateStr, "0");

                        foreach (var row in kospiRows)
                            AddSectorData(sectorSupply, sectorNameMap, sectorMarketMap, row, dTrade, "KOSPI");
                        foreach (var row in kosdaqRows)
                            AddSectorData(sectorSupply, sectorNameMap, sectorMarketMap, row, dTrade, "KOSDAQ");

                        if (actualTradingDays.Count <= 2 || actualTradingDays.Count % 5 == 0)
                            LogMsg($"  {dateStr}: 코스피 {kospiRows.Count}업종, 코스닥 {kosdaqRows.Count}업종 [{actualTradingDays.Count}/{TARGET_DAYS}]");

                        d = dTrade.AddDays(-1);
                    }

                    int nKospi = sectorMarketMap.Values.Count(v => v == "KOSPI");
                    int nKosdaq = sectorMarketMap.Values.Count(v => v == "KOSDAQ");
                    LogMsg($"  ✓ 업종수급 완료: {sectorSupply.Count}개 업종 (코스피 {nKospi}, 코스닥 {nKosdaq}), {actualTradingDays.Count}거래일");
                }
                catch (Exception ex)
                {
                    LogMsg($"  ✗ 업종수급 조회 실패: {ex.Message}");
                }

                foreach (var kv in sectorSupply)
                    kv.Value.Sort((a, b) => b.Date.CompareTo(a.Date));

                // ── 업종이름 → opt10051 복합키 역맵 (시장 포함) ──────
                var sectorNameToKey = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in sectorNameMap)
                {
                    var market = sectorMarketMap.TryGetValue(kv.Key, out var m) ? m : "KOSPI";
                    var name = NormalizeSectorName(kv.Value);
                    var key = SectorNameMapKey(market, name);
                    if (!string.IsNullOrEmpty(name) && !sectorNameToKey.ContainsKey(key))
                        sectorNameToKey[key] = kv.Key;
                }

                // ── 4. 업종 PER/PBR 조회 (opt20001) ─────────────
                LogMsg("▶ 업종 PER/PBR 조회 (opt20001)...");
                var sectorFunds = new Dictionary<string, SectorFundamental>(); // key=market|sectorCode

                // 종목들이 속한 업종 목록 수집 (시장 포함)
                var neededSectorKeys = new HashSet<string>();
                foreach (var code in codes)
                {
                    if (stockSectorMap.TryGetValue(code, out var secName))
                    {
                        var market = stockMarketMap.TryGetValue(code, out var mkt) ? mkt : (stockInfos.TryGetValue(code, out var si) ? si.Market : "KOSPI");
                        var norm = NormalizeSectorName(secName);
                        if (sectorNameToKey.TryGetValue(SectorNameMapKey(market, norm), out var secKey))
                            neededSectorKeys.Add(secKey);
                    }
                }

                foreach (var secKey in neededSectorKeys)
                {
                    ct.ThrowIfCancellationRequested();

                    string market = sectorMarketMap.TryGetValue(secKey, out var m) ? m : "KOSPI";
                    string secCode = RawSectorCode(secKey);
                    var mktType = market == "KOSDAQ" ? "1" : "0";

                    var sf = await kiwoom.GetSectorPerPbrAsync(secCode, mktType);
                    sf.SectorCode = secCode;
                    sf.SectorName = sectorNameMap.TryGetValue(secKey, out var sn) ? sn : secCode;
                    sectorFunds[secKey] = sf;

                    if (sf.AvgPer.HasValue || sf.AvgPbr.HasValue)
                        LogMsg($"  [{market}] {sf.SectorName}: PER={sf.AvgPer?.ToString("F2") ?? "-"}, PBR={sf.AvgPbr?.ToString("F2") ?? "-"}");
                }

                if (sectorFunds.Count == 0 || sectorFunds.Values.All(f => !f.AvgPer.HasValue && !f.AvgPbr.HasValue))
                    LogMsg("  ⚠ 업종 PER/PBR 데이터를 가져오지 못했습니다.");
                else
                    LogMsg($"  ✓ {sectorFunds.Count}개 업종 PER/PBR 조회 완료");

                // ── 5. 종목별 분석 ──────────────────────────────────
                LogMsg($"▶ 종목별 분석 시작 ({codes.Count}개)...");
                var results = new List<AnalysisResult>();

                for (int i = 0; i < codes.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    var code = codes[i];
                    var info = stockInfos[code];

                    Progress?.Invoke(i + 1, codes.Count, info.Name);
                    LogMsg($"  [{i + 1}/{codes.Count}] {info.Name} ({code})");

                    try
                    {
                        var fund = await kiwoom.GetFundamentalAsync(code);
                        var investors = await kiwoom.GetInvestorDataAsync(code, fromDate60);
                        var bars = await kiwoom.GetDailyBarAsync(code, fromDate60);

                        // opt10081에 시가총액 필드가 없으므로 opt10001 값으로 보정
                        if (fund != null && fund.MarketCap > 0)
                        {
                            foreach (var b in bars)
                                if (b.MarketCap <= 0) b.MarketCap = fund.MarketCap;
                        }

                        // 현재가 보정 (금액→수량 환산용 표시에 사용)
                        if (info.CurrentPrice <= 0 && bars.Count > 0)
                        {
                            var b0 = bars[0];
                            if (b0.Volume > 0 && b0.TradeAmount > 0)
                                info.CurrentPrice = b0.TradeAmount / b0.Volume;
                        }

                        // 업종수급 매칭 (시장 포함 복합키 우선)
                        List<SectorSupplyDay> ss = null;
                        string matchedSectorKey = null;
                        SectorFundamental sf = null;

                        if (!string.IsNullOrEmpty(info.SectorName))
                        {
                            var normName = NormalizeSectorName(info.SectorName);
                            var nameKey = SectorNameMapKey(info.Market, normName);
                            if (sectorNameToKey.TryGetValue(nameKey, out var sKey))
                            {
                                if (sectorSupply.TryGetValue(sKey, out ss))
                                    matchedSectorKey = sKey;
                                sectorFunds.TryGetValue(sKey, out sf);
                            }
                        }

                        if (matchedSectorKey != null)
                            info.SectorCode = RawSectorCode(matchedSectorKey);

                        LogMsg($"    업종: {info.SectorName} / {info.Market} → [{(matchedSectorKey != null ? RawSectorCode(matchedSectorKey) : "없음")}] ({(ss?.Count ?? 0)}일) PER={sf?.AvgPer?.ToString("F1") ?? "-"}");

                        var result = StockScorer.Calculate(
                            info, fund, sf, investors, bars,
                            ss ?? new List<SectorSupplyDay>(), cfg);
                        results.Add(result);

                        LogMsg($"    ✓ 총점={result.TotalScore:F1} (가치={result.ValueScore:F1} 수급={result.StockSupplyScore:F1} 업종={result.SectorSupplyScore:F1}) 거래량={result.VolTrend}");
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        results.Add(new AnalysisResult
                        { Code = code, Name = info.Name, Status = "오류", ErrorMsg = ex.Message });
                        LogMsg($"    ✗ {ex.Message}");
                    }
                }

                // ── 6. 전체 업종 수급 현황 ──────────────────────────
                LogMsg("▶ 전체 업종 수급 현황 생성...");
                var sectorKospi = new List<SectorSupplySummary>();
                var sectorKosdaq = new List<SectorSupplySummary>();

                foreach (var kv in sectorSupply)
                {
                    var recent5 = kv.Value.Take(5).ToList();
                    if (recent5.Count == 0) continue;

                    string market = sectorMarketMap.TryGetValue(kv.Key, out var mm) ? mm : "KOSPI";
                    string sName = sectorNameMap.TryGetValue(kv.Key, out var n) ? n : RawSectorCode(kv.Key);

                    var summary = new SectorSupplySummary
                    {
                        SectorCode = RawSectorCode(kv.Key),
                        SectorName = sName,
                        Market = market,
                        ForeignNet5D = recent5.Sum(dd => dd.ForeignNet),
                        InstNet5D = recent5.Sum(dd => dd.InstNet),
                    };
                    summary.TotalNet5D = summary.ForeignNet5D + summary.InstNet5D;

                    if (market == "KOSDAQ") sectorKosdaq.Add(summary);
                    else sectorKospi.Add(summary);
                }
                sectorKospi.Sort((a, b) => b.TotalNet5D.CompareTo(a.TotalNet5D));
                sectorKosdaq.Sort((a, b) => b.TotalNet5D.CompareTo(a.TotalNet5D));

                LogMsg($"  업종 현황: 코스피 {sectorKospi.Count}개, 코스닥 {sectorKosdaq.Count}개");

                results.Sort((a, b) => b.TotalScore.CompareTo(a.TotalScore));
                LogMsg($"▶ 분석 완료: {results.Count}개 종목");
                return (results, sectorKospi, sectorKosdaq);
            }
        }

        // ── opt10051 헬퍼 ─────────────────────────────────────────
        private static void AddSectorData(
            Dictionary<string, List<SectorSupplyDay>> sectorSupply,
            Dictionary<string, string> sectorNameMap,
            Dictionary<string, string> sectorMarketMap,
            SectorInvestorRow row, DateTime date, string market)
        {
            var key = SectorCompositeKey(market, row.SectorCode);

            if (!sectorSupply.ContainsKey(key))
                sectorSupply[key] = new List<SectorSupplyDay>();

            sectorSupply[key].Add(new SectorSupplyDay
            {
                Date = date,
                ForeignNet = row.ForeignNet,
                InstNet = row.InstNet,
            });

            if (!sectorNameMap.ContainsKey(key))
                sectorNameMap[key] = row.SectorName;

            if (!sectorMarketMap.ContainsKey(key))
                sectorMarketMap[key] = market;
        }

        private static string SectorCompositeKey(string market, string sectorCode)
            => $"{(market ?? "KOSPI").Trim().ToUpperInvariant()}|{(sectorCode ?? "").Trim()}";

        private static string SectorNameMapKey(string market, string normalizedSectorName)
            => $"{(market ?? "KOSPI").Trim().ToUpperInvariant()}|{normalizedSectorName ?? ""}";

        private static string RawSectorCode(string compositeKey)
        {
            if (string.IsNullOrWhiteSpace(compositeKey)) return "";
            var idx = compositeKey.IndexOf('|');
            return idx >= 0 ? compositeKey.Substring(idx + 1) : compositeKey;
        }

        // ── KOA_Functions 파싱 ────────────────────────────────────
        private static (string Market, string SectorName) ParseStockInfo(string raw)
        {
            string market = null;
            string sector = null;

            if (string.IsNullOrWhiteSpace(raw)) return (market, sector);

            foreach (var part in raw.Split(';'))
            {
                var fields = part.Split('|');
                if (fields.Length < 2) continue;

                var key = fields[0].Trim();

                if (key == "시장구분0" && fields.Length >= 2)
                {
                    var mkt = fields[1].Trim();
                    if (mkt.Contains("코스닥") || mkt == "KOSDAQ") market = "KOSDAQ";
                    else if (mkt.Contains("코스피") || mkt == "KOSPI" || mkt == "거래소") market = "KOSPI";
                }
                else if (key == "업종구분" && fields.Length >= 2)
                {
                    var name = fields[1].Trim();
                    if (!string.IsNullOrEmpty(name))
                        sector = name;
                }
            }

            return (market, sector);
        }

        private static string NormalizeSectorName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "";
            var idx = name.IndexOf('(');
            if (idx > 0) name = name.Substring(0, idx);
            return name.Trim().Replace(" ", "").Replace("·", "/");
        }

        private void LogMsg(string msg) => Log?.Invoke(msg);
    }
}
