using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using AxKHOpenAPILib;
using StockAnalyzer.Models;
using StockAnalyzer.Utils;

namespace StockAnalyzer.Api
{
    public class KiwoomClient : IDisposable
    {
        private readonly AxKHOpenAPI _ax;
        private readonly SemaphoreSlim _trLock = new SemaphoreSlim(1, 1);
        private readonly int _trDelayMs = 200;

        private TaskCompletionSource<bool> _loginTcs;
        private readonly Dictionary<string, TaskCompletionSource<TrResult>> _pending
            = new Dictionary<string, TaskCompletionSource<TrResult>>();

        private TaskCompletionSource<string> _conditionTcs;
        private TaskCompletionSource<List<string>> _conditionResultTcs;

        private int _scrSeq = 1000;
        private string NextScr() { _scrSeq = (_scrSeq % 9998) + 1; return _scrSeq.ToString("D4"); }

        public KiwoomClient(AxKHOpenAPI ax)
        {
            _ax = ax;
            _ax.OnEventConnect += OnEventConnect;
            _ax.OnReceiveTrData += OnReceiveTrData;
            _ax.OnReceiveMsg += OnReceiveMsg;
            _ax.OnReceiveConditionVer += OnReceiveConditionVer;
            _ax.OnReceiveTrCondition += OnReceiveTrCondition;
        }

        // ── 로그인 ──────────────────────────────────────────────

        public Task<bool> LoginAsync(CancellationToken ct = default)
        {
            _loginTcs = new TaskCompletionSource<bool>();
            ct.Register(() => _loginTcs.TrySetCanceled());
            _ax.CommConnect();
            return _loginTcs.Task;
        }

        public bool IsLoggedIn() => _ax.GetConnectState() == 1;
        public string GetLoginInfo(string t) => _ax.GetLoginInfo(t);

        private void OnEventConnect(object s, _DKHOpenAPIEvents_OnEventConnectEvent e)
        {
            if (e.nErrCode == 0) _loginTcs?.TrySetResult(true);
            else _loginTcs?.TrySetException(
                new Exception("로그인 오류: " + e.nErrCode));
        }

        // ── 조건검색식 ──────────────────────────────────────────

        public async Task<List<(string idx, string name)>> GetConditionListAsync()
        {
            _conditionTcs = new TaskCompletionSource<string>();
            _ax.GetConditionLoad();
            var raw = await _conditionTcs.Task;

            var result = new List<(string, string)>();
            if (string.IsNullOrEmpty(raw)) return result;

            foreach (var item in raw.Split(';'))
            {
                var parts = item.Split('^');
                if (parts.Length == 2)
                    result.Add((parts[0].Trim(), parts[1].Trim()));
            }
            return result;
        }

        private void OnReceiveConditionVer(object s, _DKHOpenAPIEvents_OnReceiveConditionVerEvent e)
        {
            if (e.lRet == 1)
                _conditionTcs?.TrySetResult(_ax.GetConditionNameList());
            else
                _conditionTcs?.TrySetException(new Exception("조건식 로드 실패"));
        }

        public async Task<List<string>> GetConditionCodesAsync(string idx, string name)
        {
            _conditionResultTcs = new TaskCompletionSource<List<string>>();
            _ax.SendCondition(NextScr(), name, int.Parse(idx), 0);
            return await _conditionResultTcs.Task;
        }

        private void OnReceiveTrCondition(object s, _DKHOpenAPIEvents_OnReceiveTrConditionEvent e)
        {
            var codes = new List<string>();
            if (!string.IsNullOrEmpty(e.strCodeList))
                foreach (var c in e.strCodeList.Split(';'))
                    if (c.Trim().Length == 6) codes.Add(c.Trim());
            _conditionResultTcs?.TrySetResult(codes);
        }

        // ── TR 요청 ─────────────────────────────────────────────

        private async Task<TrResult> RequestAsync(
            string rqName, string trCode, int prevNext,
            Dictionary<string, string> inputs, int timeoutMs = 15000)
        {
            await _trLock.WaitAsync();
            try
            {
                await Task.Delay(_trDelayMs);
                var tcs = new TaskCompletionSource<TrResult>();
                _pending[rqName] = tcs;

                foreach (var kv in inputs)
                    _ax.SetInputValue(kv.Key, kv.Value);

                int ret = _ax.CommRqData(rqName, trCode, prevNext, NextScr());
                if (ret != 0)
                {
                    _pending.Remove(rqName);
                    throw new Exception($"CommRqData 오류({ret}): {rqName}/{trCode}");
                }

                var cts = new CancellationTokenSource(timeoutMs);
                cts.Token.Register(() => tcs.TrySetCanceled());
                return await tcs.Task;
            }
            finally { _trLock.Release(); }
        }

        private void OnReceiveTrData(object s, _DKHOpenAPIEvents_OnReceiveTrDataEvent e)
        {
            if (!_pending.TryGetValue(e.sRQName, out var tcs)) return;
            _pending.Remove(e.sRQName);
            tcs.TrySetResult(new TrResult(_ax, e.sTrCode, e.sRQName, e.sPrevNext));
        }

        private void OnReceiveMsg(object s, _DKHOpenAPIEvents_OnReceiveMsgEvent e)
            => System.Diagnostics.Debug.WriteLine($"[Msg] {e.sRQName} | {e.sMsg}");

        // ── opt10001 기업가치 ────────────────────────────────────

        public async Task<FundamentalData> GetFundamentalAsync(string code)
        {
            var r = await RequestAsync("주식기본정보", "opt10001", 0,
                new Dictionary<string, string> { ["종목코드"] = code });

            return new FundamentalData
            {
                Code = code,
                Per = r.GetDouble("주식기본정보", "PER"),
                Pbr = r.GetDouble("주식기본정보", "PBR"),
                Roe = r.GetDouble("주식기본정보", "ROE"),
                Eps = r.GetDouble("주식기본정보", "EPS"),
                Bps = r.GetDouble("주식기본정보", "BPS"),
                MarketCap = r.GetDouble2("주식기본정보", "시가총액") * 100_000_000, // 억→원
            };
        }

        // ── opt10059 투자자별 순매수 (수량 모드, 60일) ──────────

        public async Task<List<InvestorDay>> GetInvestorDataAsync(
            string code, DateTime fromDate)
        {
            int prevNext = 0;
            var result = new List<InvestorDay>();
            var cutoff = fromDate.Date;

            while (true)
            {
                var r = await RequestAsync("종목투자자", "opt10059", prevNext,
                    new Dictionary<string, string>
                    {
                        ["일자"] = TradingDayHelper.ToApiDate(DateTime.Today),
                        ["종목코드"] = code,
                        ["금액수량구분"] = "1",   // 금액(백만원) — 수량모드 미지원 이슈로 금액 사용
                        ["매매구분"] = "0",
                        ["단위구분"] = "1",
                    });

                int cnt = r.RepeatCount("종목별투자자기관별");
                bool done = false;
                for (int i = 0; i < cnt; i++)
                {
                    if (!TryParseDate(r.GetString("종목별투자자기관별", "일자", i), out var d)) continue;
                    if (d.Date < cutoff) { done = true; break; }
                    result.Add(new InvestorDay
                    {
                        Date = d,
                        ForeignNet = r.GetLong("종목별투자자기관별", "외국인투자자", i) * 1_000_000, // 백만원→원
                        InstNet = r.GetLong("종목별투자자기관별", "기관계", i) * 1_000_000,       // 백만원→원
                        Volume = Math.Abs(r.GetLong("종목별투자자기관별", "누적거래량", i)),
                    });
                }

                if (done || !r.HasNext) break;
                prevNext = r.HasNext ? 2 : 0;
            }
            return result;
        }

        // ── opt10081 일봉 (거래대금/시가총액/거래량, 60일) ──────

        public async Task<List<DailyBar>> GetDailyBarAsync(
            string code, DateTime fromDate)
        {
            int prevNext = 0;
            var result = new List<DailyBar>();
            var cutoff = fromDate.Date;

            while (true)
            {
                var r = await RequestAsync("주식일봉", "opt10081", prevNext,
                    new Dictionary<string, string>
                    {
                        ["종목코드"] = code,
                        ["기준일자"] = TradingDayHelper.ToApiDate(DateTime.Today),
                        ["수정주가구분"] = "1",
                    });

                int cnt = r.RepeatCount("주식일봉차트조회");
                bool done = false;
                for (int i = 0; i < cnt; i++)
                {
                    if (!TryParseDate(r.GetString("주식일봉차트조회", "일자", i), out var d)) continue;
                    if (d.Date < cutoff) { done = true; break; }
                    result.Add(new DailyBar
                    {
                        Date = d,
                        MarketCap = r.GetDouble2("주식일봉차트조회", "시가총액", i) * 1_000_000,
                        TradeAmount = r.GetDouble2("주식일봉차트조회", "거래대금", i) * 1_000_000,
                        Volume = Math.Abs(r.GetLong("주식일봉차트조회", "거래량", i)),
                    });
                }
                if (done || !r.HasNext) break;
                prevNext = r.HasNext ? 2 : 0;
            }
            return result;
        }

        // ── opt10051 업종별투자자순매수 ─────────────────────────

        public async Task<List<SectorInvestorRow>> GetSectorInvestorAsync(
            string marketCode, string date, string amtQtyType = "0")
        {
            var r = await RequestAsync("업종별투자자", "OPT10051", 0,
                new Dictionary<string, string>
                {
                    ["시장구분"] = marketCode,
                    ["금액수량구분"] = amtQtyType,
                    ["기준일자"] = date,
                    ["거래소구분"] = "",
                });

            var result = new List<SectorInvestorRow>();
            int cnt = r.RepeatCount("업종별순매수");
            for (int i = 0; i < cnt; i++)
            {
                var sectorCode = r.GetString("업종별순매수", "업종코드", i);
                if (string.IsNullOrWhiteSpace(sectorCode)) continue;

                result.Add(new SectorInvestorRow
                {
                    SectorCode = sectorCode.Trim(),
                    SectorName = r.GetString("업종별순매수", "업종명", i),
                    ForeignNet = r.GetDouble2("업종별순매수", "외국인순매수", i) * 1_000_000,
                    InstNet = r.GetDouble2("업종별순매수", "기관계순매수", i) * 1_000_000,
                });
            }
            return result;
        }

        // ── opt20001 업종현재가 → PER/PBR ──────────────────────

        /// <summary>
        /// 업종 PER/PBR을 조회합니다.
        /// 업종코드: opt10051에서 사용하는 코드 (예: "001"=종합, "027"=화학)
        /// 시장구분: "0"=코스피, "1"=코스닥
        /// </summary>
        public async Task<SectorFundamental> GetSectorPerPbrAsync(
            string sectorCode, string marketType)
        {
            try
            {
                var r = await RequestAsync("업종현재가_" + sectorCode, "opt20001", 0,
                    new Dictionary<string, string>
                    {
                        ["시장구분"] = marketType,
                        ["업종코드"] = sectorCode,
                    });

                // opt20001 비반복 필드에서 PER 추출 시도
                var per = r.GetDouble("업종현재가", "PER");
                if (per == null) per = r.GetDouble("업종현재가", "업종PER");
                var pbr = r.GetDouble("업종현재가", "PBR");
                if (pbr == null) pbr = r.GetDouble("업종현재가", "업종PBR");

                return new SectorFundamental
                {
                    SectorCode = sectorCode,
                    AvgPer = per,
                    AvgPbr = pbr,
                };
            }
            catch
            {
                return new SectorFundamental { SectorCode = sectorCode };
            }
        }

        /// <summary>opt20001 필드 덤프 (진단용)</summary>
        public async Task<string> DumpSectorFieldsAsync(string sectorCode, string marketType)
        {
            try
            {
                var r = await RequestAsync("업종현재가_덤프", "opt20001", 0,
                    new Dictionary<string, string>
                    {
                        ["시장구분"] = marketType,
                        ["업종코드"] = sectorCode,
                    });

                var fields = new[] {
                    "현재가","전일대비","거래량","거래대금","PER","PBR",
                    "업종PER","업종PBR","시가총액","등락률",
                    "상한","상승","보합","하한","하락",
                    "업종코드","업종명","시장구분"
                };
                var nonRepeat = r.DumpFields("업종현재가", fields);
                int cnt = r.RepeatCount("업종현재가");
                return $"[opt20001] 비반복: {nonRepeat}, 반복={cnt}건";
            }
            catch (Exception ex)
            {
                return $"[opt20001] 오류: {ex.Message}";
            }
        }

        // ── 진단용: 필드명 덤프 ──────────────────────────────────

        public async Task<string> DumpFieldsAsync(string code)
        {
            var sb = new System.Text.StringBuilder();

            var r01 = await RequestAsync("주식기본정보_진단", "opt10001", 0,
                new Dictionary<string, string> { ["종목코드"] = code });
            sb.AppendLine("── opt10001 필드 덤프 ──");
            var fields01 = new[] {
                "종목코드","종목명","결산월","액면가","자본금","상장주식","신용비율",
                "시가총액","외인소진률","PER","EPS","ROE","PBR","EV","BPS",
                "매출액","영업이익","당기순이익","현재가","전일대비","거래량","유통주식","유통비율"
            };
            sb.AppendLine("  " + r01.DumpFields("주식기본정보", fields01, 0));

            return sb.ToString();
        }

        private static bool TryParseDate(string s, out DateTime d)
        {
            d = default;
            s = s?.Trim().Replace("-", "").Replace("/", "");
            return s?.Length == 8 &&
                   DateTime.TryParseExact(s, "yyyyMMdd",
                       System.Globalization.CultureInfo.InvariantCulture,
                       System.Globalization.DateTimeStyles.None, out d);
        }

        public void Dispose()
        {
            _ax.OnEventConnect -= OnEventConnect;
            _ax.OnReceiveTrData -= OnReceiveTrData;
            _ax.OnReceiveMsg -= OnReceiveMsg;
            _ax.OnReceiveConditionVer -= OnReceiveConditionVer;
            _ax.OnReceiveTrCondition -= OnReceiveTrCondition;
            _trLock.Dispose();
        }
    }

    // ── TR 응답 래퍼 ─────────────────────────────────────────────

    public class TrResult
    {
        private readonly AxKHOpenAPI _ax;
        private readonly string _tr;
        public string PrevNext { get; }
        public bool HasNext => PrevNext == "2";

        public TrResult(AxKHOpenAPI ax, string tr, string rq, string pn)
        { _ax = ax; _tr = tr; PrevNext = pn; }

        public string GetString(string rec, string fld, int idx = 0)
            => _ax.GetCommData(_tr, rec, idx, fld).Trim();

        public double? GetDouble(string rec, string fld, int idx = 0)
        {
            var s = GetString(rec, fld, idx).Replace(",", "").Replace("+", "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return null;
            return double.TryParse(s, out var v) ? v : (double?)null;
        }

        public double GetDouble2(string rec, string fld, int idx = 0)
        {
            var s = GetString(rec, fld, idx).Replace(",", "").Replace("+", "");
            return double.TryParse(s, out var v) ? v : 0;
        }

        public long GetLong(string rec, string fld, int idx = 0)
        {
            var s = GetString(rec, fld, idx).Replace(",", "").Replace("+", "");
            return long.TryParse(s, out var v) ? v : 0;
        }

        public int RepeatCount(string rec) => _ax.GetRepeatCnt(_tr, rec);

        public string DumpFields(string rec, string[] fields, int idx = 0)
        {
            var parts = new List<string>();
            foreach (var f in fields)
            {
                var v = GetString(rec, f, idx);
                if (!string.IsNullOrEmpty(v))
                    parts.Add($"{f}=[{v}]");
            }
            return string.Join(", ", parts);
        }
    }
}