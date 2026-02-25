using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using StockAnalyzer.Models;

namespace StockAnalyzer.Api
{
    /// <summary>
    /// 컨센서스 데이터 스크래핑 클라이언트
    /// Source: navercomp.wisereport.co.kr (FnGuide)
    /// 기업현황 페이지에서 투자의견 컨센서스 + 증권사별 목표가 파싱
    /// </summary>
    public class ConsensusClient : IDisposable
    {
        private readonly HttpClient _http;
        private readonly Dictionary<string, ConsensusData> _cache
            = new Dictionary<string, ConsensusData>();
        private static readonly TimeSpan CacheTtl = TimeSpan.FromMinutes(30);

        public ConsensusClient()
        {
            _http = new HttpClient();
            _http.DefaultRequestHeaders.Add("User-Agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
                "AppleWebKit/537.36 (KHTML, like Gecko) " +
                "Chrome/120.0.0.0 Safari/537.36");
            _http.DefaultRequestHeaders.Add("Accept",
                "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
            _http.DefaultRequestHeaders.Add("Accept-Language", "ko-KR,ko;q=0.9,en;q=0.5");
            _http.DefaultRequestHeaders.Add("Referer", "https://finance.naver.com/");
            _http.Timeout = TimeSpan.FromSeconds(15);
        }

        /// <summary>
        /// 종목코드로 컨센서스 조회 (캐시 있으면 캐시 반환)
        /// </summary>
        public async Task<ConsensusData> GetConsensusAsync(string code, CancellationToken ct = default)
        {
            if (string.IsNullOrEmpty(code)) return null;

            // 캐시 확인
            if (_cache.TryGetValue(code, out var cached)
                && (DateTime.Now - cached.FetchedAt) < CacheTtl)
                return cached;

            ConsensusData result = null;

            try
            {
                result = await FetchWiseReportAsync(code, ct);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Consensus] WiseReport 실패({code}): {ex.Message}");
            }

            if (result != null)
            {
                result.Code = code;
                result.FetchedAt = DateTime.Now;
                _cache[code] = result;
            }

            return result;
        }

        // ────────────────────────────────────────────────────────
        //  navercomp.wisereport.co.kr 기업현황 스크래핑
        //  투자의견 컨센서스 요약 + 증권사별 목표가 테이블
        // ────────────────────────────────────────────────────────

        private async Task<ConsensusData> FetchWiseReportAsync(string code, CancellationToken ct)
        {
            var url = $"https://navercomp.wisereport.co.kr/v2/company/c1010001.aspx?cmp_cd={code}&target=finsum_more";
            string html;
            using (var req = new HttpRequestMessage(HttpMethod.Get, url))
            using (var resp = await _http.SendAsync(req, ct))
            {
                resp.EnsureSuccessStatusCode();
                html = await resp.Content.ReadAsStringAsync();
            }

            var data = new ConsensusData { Source = "WiseReport" };

            // ═══════════════════════════════════════════════════════
            //  현재가 파싱 (요약정보 섹션)
            //  "주가/전일대비/수익률 | 181,200원 / +2,600원"
            // ═══════════════════════════════════════════════════════
            var priceMatch = Regex.Match(html,
                @"주가/전일대비.*?([\d,]+)\s*원",
                RegexOptions.Singleline);
            if (priceMatch.Success)
                data.CurrentPrice = ParseNum(priceMatch.Groups[1].Value);

            // 전일종가 fallback
            if (!data.CurrentPrice.HasValue)
            {
                var closeMatch = Regex.Match(html,
                    @"전일종가\s*\**\s*([\d,]+)",
                    RegexOptions.Singleline);
                if (closeMatch.Success)
                    data.CurrentPrice = ParseNum(closeMatch.Groups[1].Value);
            }

            // ═══════════════════════════════════════════════════════
            //  투자의견 컨센서스 요약 테이블
            //  "투자의견 컨센서스" 섹션
            //  투자의견(점수) | 목표주가(원) | EPS(원) | PER(배) | 추정기관수
            // ═══════════════════════════════════════════════════════
            var csSection = Regex.Match(html,
                @"투자의견\s*컨센서스.*?</table>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            if (csSection.Success)
            {
                var sec = csSection.Value;

                // 투자의견 점수 (1~5), 패턴: "값:4.00" 또는 "**4.00**"
                var ratingMatch = Regex.Match(sec,
                    @"값\s*[:\s]*([\d.]+)",
                    RegexOptions.Singleline);
                if (ratingMatch.Success)
                {
                    double rating;
                    if (double.TryParse(ratingMatch.Groups[1].Value, NumberStyles.Any,
                        CultureInfo.InvariantCulture, out rating))
                    {
                        data.Opinion = RatingToOpinion(rating);
                    }
                }

                // 테이블 데이터 셀에서 숫자 추출
                var cells = Regex.Matches(sec,
                    @"<td[^>]*>\s*([\d,.]+)\s*</td>",
                    RegexOptions.Singleline | RegexOptions.IgnoreCase);

                if (cells.Count >= 4)
                {
                    // [0]=목표주가, [1]=EPS, [2]=PER, [3]=추정기관수
                    data.TargetPrice = ParseNum(cells[0].Groups[1].Value);
                    data.ConsensusEps = ParseNum(cells[1].Groups[1].Value);
                    data.ConsensusPer = ParseDbl(cells[2].Groups[1].Value);
                    int cnt;
                    if (int.TryParse(cells[3].Groups[1].Value.Replace(",", ""), out cnt))
                        data.AnalystCount = cnt;
                }
                else
                {
                    // 대체: bold 또는 strong 태그 안 숫자
                    var boldNums = Regex.Matches(sec,
                        @"<(?:strong|b|td)[^>]*>\s*\**\s*([\d,.]+)\s*\**\s*</(?:strong|b|td)>",
                        RegexOptions.Singleline | RegexOptions.IgnoreCase);
                    var nums = new List<string>();
                    foreach (Match m in boldNums)
                        nums.Add(m.Groups[1].Value);

                    if (nums.Count >= 5)
                    {
                        // [0]=rating, [1]=목표주가, [2]=EPS, [3]=PER, [4]=추정기관수
                        data.TargetPrice = ParseNum(nums[1]);
                        data.ConsensusEps = ParseNum(nums[2]);
                        data.ConsensusPer = ParseDbl(nums[3]);
                        int cnt;
                        if (int.TryParse(nums[4].Replace(",", ""), out cnt))
                            data.AnalystCount = cnt;
                    }
                }
            }

            // ═══════════════════════════════════════════════════════
            //  증권사별 목표가 테이블 파싱
            //  | 제공처 | 최종일자 | 목표가 | 직전목표가 | 변동률(%) | 투자의견 |
            // ═══════════════════════════════════════════════════════
            var brokerSection = Regex.Match(html,
                @"제공처별\s*투자의견.*?</table>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            if (!brokerSection.Success)
            {
                brokerSection = Regex.Match(html,
                    @"<table[^>]*>(?=.*?제공처)(?=.*?목표가).*?</table>",
                    RegexOptions.Singleline | RegexOptions.IgnoreCase);
            }

            var targetPrices = new List<double>();
            var reportDates = new List<string>();

            if (brokerSection.Success)
            {
                var rows = Regex.Matches(brokerSection.Value,
                    @"<tr[^>]*>(.*?)</tr>",
                    RegexOptions.Singleline | RegexOptions.IgnoreCase);

                foreach (Match row in rows)
                {
                    var tds = Regex.Matches(row.Groups[1].Value,
                        @"<td[^>]*>(.*?)</td>",
                        RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    if (tds.Count < 6) continue;

                    // tds[1] = 최종일자 (26/02/13)
                    var dateRaw = StripTags(tds[1].Groups[1].Value).Trim();
                    // tds[2] = 목표가 (240,000)
                    var priceRaw = StripTags(tds[2].Groups[1].Value).Trim();

                    if (Regex.IsMatch(dateRaw, @"\d{2}/\d{2}/\d{2}"))
                        reportDates.Add(dateRaw);

                    var tp = ParseNum(priceRaw);
                    if (tp.HasValue && tp.Value > 0)
                        targetPrices.Add(tp.Value);
                }
            }

            // ── 리포트 결과 반영 (최솟값·최댓값 제외한 중간 범위) ──
            if (targetPrices.Count > 0)
            {
                if (data.AnalystCount == 0)
                    data.AnalystCount = targetPrices.Count;

                var sorted = targetPrices.OrderBy(p => p).ToList();

                if (sorted.Count >= 3)
                {
                    // 최솟값·최댓값 각 1건 제외
                    var mid = sorted.Skip(1).Take(sorted.Count - 2).ToList();
                    data.TargetPriceMin = mid.Min();
                    data.TargetPriceMax = mid.Max();
                    if (!data.TargetPrice.HasValue)
                        data.TargetPrice = Math.Round(mid.Average());
                }
                else
                {
                    // 2건 이하면 그대로
                    data.TargetPriceMin = sorted.First();
                    data.TargetPriceMax = sorted.Last();
                    if (!data.TargetPrice.HasValue)
                        data.TargetPrice = Math.Round(sorted.Average());
                }
            }

            if (reportDates.Count > 0)
            {
                data.LatestReportDate = FormatDate(
                    reportDates.OrderByDescending(d => d).First());
            }

            // ── 괴리율 계산 (평균 목표가 기준) ──
            if (data.TargetPrice.HasValue && data.CurrentPrice.HasValue && data.CurrentPrice.Value > 0)
                data.DeviationPct = (data.TargetPrice.Value - data.CurrentPrice.Value)
                                    / data.CurrentPrice.Value * 100.0;

            System.Diagnostics.Debug.WriteLine(
                $"[Consensus] WiseReport {code}: 의견={data.Opinion}, " +
                $"평균={data.TargetPrice:N0}, 중간범위={data.TargetPriceMin:N0}~{data.TargetPriceMax:N0}, " +
                $"증권사수={data.AnalystCount}, 최신={data.LatestReportDate}, 현재가={data.CurrentPrice:N0}");

            return data;
        }

        // ────────────────────────────────────────────────────────
        //  유틸
        // ────────────────────────────────────────────────────────

        /// <summary>투자의견 점수 → 한글 (5=강력매수 ... 1=강력매도)</summary>
        private static string RatingToOpinion(double rating)
        {
            if (rating >= 4.21) return "강력매수";
            if (rating >= 3.41) return "매수";
            if (rating >= 2.61) return "보유";
            if (rating >= 1.81) return "매도";
            return "강력매도";
        }

        private static string StripTags(string html)
        {
            return Regex.Replace(html ?? "", @"<[^>]+>", "").Trim();
        }

        /// <summary>"26/02/13" → "2026.02.13"</summary>
        private static string FormatDate(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return raw;
            var m = Regex.Match(raw, @"(\d{2})/(\d{2})/(\d{2})");
            if (!m.Success) return raw;
            return $"20{m.Groups[1].Value}.{m.Groups[2].Value}.{m.Groups[3].Value}";
        }

        private static double? ParseNum(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            s = s.Replace(",", "").Replace("*", "").Replace("원", "").Replace("배", "").Trim();
            double v;
            return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v) ? v : (double?)null;
        }

        private static double? ParseDbl(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            s = s.Replace(",", "").Replace("*", "").Trim();
            double v;
            return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v) ? v : (double?)null;
        }

        public void ClearCache() => _cache.Clear();
        public void Dispose() => _http?.Dispose();
    }
}