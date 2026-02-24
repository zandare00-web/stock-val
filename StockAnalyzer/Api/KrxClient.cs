using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using StockAnalyzer.Models;
using StockAnalyzer.Utils;

namespace StockAnalyzer.Api
{
    /// <summary>
    /// KRX 공식 Open API 클라이언트
    /// https://openapi.krx.co.kr 에서 발급받은 AUTH_KEY 사용
    /// Base URL: https://data-dbg.krx.co.kr/svc/apis/
    /// 
    /// 필요한 API 이용신청 목록:
    ///   - 유가증권 종목기본정보 (stk_isu_base_info)
    ///   - 코스닥 종목기본정보   (ksq_isu_base_info)
    ///   - 유가증권 일별매매정보 (stk_bydd_trd)
    ///   - 코스닥 일별매매정보   (ksq_bydd_trd)
    /// </summary>
    public class KrxClient : IDisposable
    {
        private readonly HttpClient _http;
        private readonly string _authKey;

        private const string API_BASE = "https://data-dbg.krx.co.kr/svc/apis";

        // 주식 엔드포인트
        private const string STK_DAILY_TRADE = "/sto/stk_bydd_trd";
        private const string STK_BASE_INFO   = "/sto/stk_isu_base_info";
        private const string KSQ_DAILY_TRADE = "/sto/ksq_bydd_trd";
        private const string KSQ_BASE_INFO   = "/sto/ksq_isu_base_info";

        public KrxClient(string authKey)
        {
            _authKey = authKey?.Trim()
                ?? throw new ArgumentNullException(nameof(authKey),
                    "KRX API 인증키가 설정되지 않았습니다. 설정에서 입력해주세요.");
            _http = new HttpClient();
            _http.DefaultRequestHeaders.Add("AUTH_KEY", _authKey);
            _http.DefaultRequestHeaders.Add("Accept", "application/json");
            _http.Timeout = TimeSpan.FromSeconds(30);
        }

        // ── API 호출 공통 ────────────────────────────────────────

        private async Task<JArray> PostAsync(string endpoint, Dictionary<string, string> parms)
        {
            var url = API_BASE + endpoint;
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(parms);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var resp = await _http.PostAsync(url, content);

            if ((int)resp.StatusCode == 401)
                throw new Exception(
                    "KRX API 인증 실패(401).\n\n" +
                    "1) AUTH_KEY가 올바른지 확인하세요.\n" +
                    "2) openapi.krx.co.kr → 마이페이지 → 이용현황에서\n" +
                    "   해당 API 서비스가 '승인' 상태인지 확인하세요.\n" +
                    "3) 필요한 서비스: 유가증권/코스닥 종목기본정보, 일별매매정보");

            resp.EnsureSuccessStatusCode();
            var body = await resp.Content.ReadAsStringAsync();
            var obj = JObject.Parse(body);

            var block = obj["OutBlock_1"] as JArray;
            return block ?? new JArray();
        }

        // ── 전체 종목 리스트 (캐시) ──────────────────────────────
        private List<StockInfo> _stockCache;
        private Dictionary<string, StockDailyTrade> _dailyTradeCache;

        public async Task<List<StockInfo>> GetAllStocksAsync()
        {
            if (_stockCache != null) return _stockCache;

            var result = new List<StockInfo>();
            var latestDate = GetLatestTradingDate();

            // KOSPI 종목기본정보
            try
            {
                var kospiInfo = await PostAsync(STK_BASE_INFO,
                    new Dictionary<string, string> { ["basDd"] = latestDate });
                foreach (var item in kospiInfo)
                {
                    var si = ParseStockInfo(item, "KOSPI");
                    if (si != null) result.Add(si);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[KRX] KOSPI 기본정보 조회 실패: {ex.Message}");
            }

            await Task.Delay(300);

            // KOSDAQ 종목기본정보
            try
            {
                var kosdaqInfo = await PostAsync(KSQ_BASE_INFO,
                    new Dictionary<string, string> { ["basDd"] = latestDate });
                foreach (var item in kosdaqInfo)
                {
                    var si = ParseStockInfo(item, "KOSDAQ");
                    if (si != null) result.Add(si);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[KRX] KOSDAQ 기본정보 조회 실패: {ex.Message}");
            }

            _stockCache = result;

            // 일별매매정보도 함께 캐시 (시세 데이터)
            await CacheDailyTradeAsync(latestDate);

            return _stockCache;
        }

        /// <summary>일별매매정보 캐시 (시가총액, 거래대금 등)</summary>
        private async Task CacheDailyTradeAsync(string date)
        {
            _dailyTradeCache = new Dictionary<string, StockDailyTrade>();

            try
            {
                var kospiTrd = await PostAsync(STK_DAILY_TRADE,
                    new Dictionary<string, string> { ["basDd"] = date });
                foreach (var item in kospiTrd)
                    ParseDailyTrade(item);
            }
            catch { }

            await Task.Delay(300);

            try
            {
                var kosdaqTrd = await PostAsync(KSQ_DAILY_TRADE,
                    new Dictionary<string, string> { ["basDd"] = date });
                foreach (var item in kosdaqTrd)
                    ParseDailyTrade(item);
            }
            catch { }

            // 종가 반영
            if (_stockCache != null)
            {
                foreach (var stock in _stockCache)
                {
                    if (_dailyTradeCache.TryGetValue(stock.Code, out var trd))
                        stock.CurrentPrice = trd.ClosePrice;
                }
            }
        }

        private void ParseDailyTrade(JToken item)
        {
            var code = ExtractCode(item);
            if (string.IsNullOrEmpty(code)) return;

            _dailyTradeCache[code] = new StockDailyTrade
            {
                Code       = code,
                ClosePrice = ParseDouble((item["TDD_CLSPRC"] ?? "0").ToString()),
                Volume     = (long)ParseDouble((item["ACC_TRDVOL"] ?? "0").ToString()),
                TradeValue = ParseDouble((item["ACC_TRDVAL"] ?? "0").ToString()),
                MarketCap  = ParseDouble((item["MKTCAP"] ?? "0").ToString()),
                ListShares = (long)ParseDouble((item["LIST_SHRS"] ?? "0").ToString()),
            };
        }

        private static StockInfo ParseStockInfo(JToken item, string market)
        {
            var code = ExtractCode(item);
            if (string.IsNullOrEmpty(code) || code.Length != 6) return null;

            return new StockInfo
            {
                Code         = code,
                Name         = (item["ISU_NM"] ?? item["ISU_ABBRV"] ?? code).ToString().Trim(),
                Market       = market,
                SectorCode   = (item["IDX_IND_CD"] ?? item["IND_TP_CD"]
                               ?? item["MKT_TP_NM"] ?? "").ToString().Trim(),
                SectorName   = (item["IDX_IND_NM"] ?? item["IND_TP_NM"] ?? "").ToString().Trim(),
                CurrentPrice = ParseDouble((item["TDD_CLSPRC"] ?? "0").ToString()),
            };
        }

        private static string ExtractCode(JToken item)
        {
            // ISU_SRT_CD(단축코드 6자리) 우선, 없으면 ISU_CD에서 추출
            var code = (item["ISU_SRT_CD"] ?? "").ToString().Trim();
            if (code.Length == 6) return code;

            code = (item["ISU_CD"] ?? "").ToString().Trim();
            if (code.Length >= 6)
            {
                // ISIN 코드(12자리)에서 뒤 6자리 추출 등
                var digits = new string(code.Where(char.IsDigit).ToArray());
                if (digits.Length >= 6) return digits.Substring(0, 6);
            }
            return code;
        }

        /// <summary>종목코드로 StockInfo 반환</summary>
        public async Task<StockInfo> GetStockInfoAsync(string code)
        {
            var all = await GetAllStocksAsync();
            return all.FirstOrDefault(s => s.Code == code);
        }

        // ── 업종 평균 PER/PBR ────────────────────────────────────
        private Dictionary<string, SectorFundamental> _sectorFundCache;

        public async Task<SectorFundamental> GetSectorFundamentalAsync(
            string sectorCode, string market)
        {
            if (_sectorFundCache == null)
                _sectorFundCache = await CalcSectorFundamentalsAsync();

            _sectorFundCache.TryGetValue(sectorCode, out var result);
            return result ?? new SectorFundamental { SectorCode = sectorCode };
        }

        /// <summary>
        /// 전종목 기본정보에서 업종별 그룹핑하여 기본 업종 정보 구축.
        /// PER/PBR은 키움 opt10001에서 가져온 데이터로 보완.
        /// </summary>
        private async Task<Dictionary<string, SectorFundamental>> CalcSectorFundamentalsAsync()
        {
            var result = new Dictionary<string, SectorFundamental>();
            var allStocks = await GetAllStocksAsync();

            var groups = allStocks
                .Where(s => !string.IsNullOrEmpty(s.SectorCode))
                .GroupBy(s => s.SectorCode);

            foreach (var g in groups)
            {
                result[g.Key] = new SectorFundamental
                {
                    SectorCode = g.Key,
                    SectorName = g.First().SectorName,
                    AvgPer     = null,  // 키움에서 보완
                    AvgPbr     = null,
                };
            }

            return result;
        }

        // ── 업종 수급 ────────────────────────────────────────────
        // KRX Open API의 일별매매정보에는 투자자별 순매수가 없으므로
        // 키움 데이터로 대체. 여기서는 빈 데이터 반환.
        public Task<List<SectorSupplyDay>> GetSectorSupplyAsync(
            string sectorCode, string market, List<DateTime> tradingDays)
        {
            return Task.FromResult(new List<SectorSupplyDay>());
        }

        // ── 전체 업종 수급 현황 (3단 상단) ───────────────────────
        // 일별매매정보 기반 업종별 거래대금 순위
        public async Task<List<SectorSupplySummary>> GetAllSectorSupplyAsync(
            string market, List<DateTime> tradingDays5D)
        {
            var allStocks = await GetAllStocksAsync();
            var marketStocks = allStocks.Where(s => s.Market == market).ToList();

            var endpoint = market == "KOSDAQ" ? KSQ_DAILY_TRADE : STK_DAILY_TRADE;
            var sectorMap = new Dictionary<string, SectorSupplySummary>();

            // 최근 1일분 거래대금으로 업종 순위 (API 호출량 절약)
            try
            {
                var dateStr = TradingDayHelper.ToApiDate(tradingDays5D[0]);
                var data = await PostAsync(endpoint,
                    new Dictionary<string, string> { ["basDd"] = dateStr });

                foreach (var item in data)
                {
                    var code = ExtractCode(item);
                    var stockInfo = marketStocks.FirstOrDefault(s => s.Code == code);
                    if (stockInfo == null || string.IsNullOrEmpty(stockInfo.SectorCode))
                        continue;

                    if (!sectorMap.TryGetValue(stockInfo.SectorCode, out var summary))
                    {
                        summary = new SectorSupplySummary
                        {
                            SectorCode = stockInfo.SectorCode,
                            SectorName = stockInfo.SectorName,
                            Market     = market,
                        };
                        sectorMap[stockInfo.SectorCode] = summary;
                    }

                    var trdVal = ParseDouble((item["ACC_TRDVAL"] ?? "0").ToString());
                    summary.TotalNet5D += trdVal;
                }
            }
            catch { }

            var result = new List<SectorSupplySummary>(sectorMap.Values);
            result.Sort((a, b) => b.TotalNet5D.CompareTo(a.TotalNet5D));
            return result;
        }

        // ── 유틸 ────────────────────────────────────────────────

        private static string GetLatestTradingDate()
        {
            // KRX Open API: 전일 데이터가 익일 08시에 업데이트
            var d = DateTime.Today;
            if (DateTime.Now.Hour < 9) d = d.AddDays(-1); // 오전 9시 전이면 전일
            d = d.AddDays(-1); // 기본적으로 전일 데이터

            while (!TradingDayHelper.IsTradingDay(d))
                d = d.AddDays(-1);
            return d.ToString("yyyyMMdd");
        }

        private static double ParseDouble(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            s = s.Replace(",", "").Replace("+", "").Trim();
            return double.TryParse(s, out var v) ? v : 0;
        }

        public void Dispose() => _http.Dispose();
    }

    /// <summary>일별매매정보 캐시용</summary>
    internal class StockDailyTrade
    {
        public string Code       { get; set; }
        public double ClosePrice { get; set; }
        public long   Volume     { get; set; }
        public double TradeValue { get; set; }
        public double MarketCap  { get; set; }
        public long   ListShares { get; set; }
    }
}
