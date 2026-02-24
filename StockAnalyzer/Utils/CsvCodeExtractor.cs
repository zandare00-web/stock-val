using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using StockAnalyzer.Models;

namespace StockAnalyzer.Utils
{
    public static class CsvCodeExtractor
    {
        private static readonly Regex _codePattern = new Regex(@"^'?(\d{6})$");

        /// <summary>
        /// CSV에서 종목코드(6자리 숫자)만 자동으로 탐지해서 추출
        /// 컬럼 순서/형식이 바뀌어도 동작
        /// </summary>
        public static List<string> Extract(string path)
        {
            var encoding = DetectEncoding(path);
            var lines = File.ReadAllLines(path, encoding);

            if (lines.Length < 2)
                throw new InvalidOperationException("CSV 파일에 데이터가 없습니다.");

            // 헤더에서 종목코드 컬럼 인덱스 찾기
            var headers = SplitCsv(lines[0]);
            int codeIdx = FindCodeColumnIndex(headers, lines);

            if (codeIdx < 0)
                throw new InvalidOperationException("종목코드 컬럼을 찾을 수 없습니다.");

            var codes = new List<string>();
            var seen  = new HashSet<string>();

            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var cols = SplitCsv(lines[i]);
                if (codeIdx >= cols.Length) continue;

                var raw  = cols[codeIdx].Trim().TrimStart('\'');
                var code = raw.PadLeft(6, '0');

                if (IsValidCode(code) && seen.Add(code))
                    codes.Add(code);
            }

            return codes;
        }

        // 헤더명 또는 데이터 패턴으로 종목코드 컬럼 탐지
        private static int FindCodeColumnIndex(string[] headers, string[] lines)
        {
            // 1차: 헤더명으로 찾기
            var codeKeywords = new[] { "종목코드", "code", "ticker", "코드" };
            for (int c = 0; c < headers.Length; c++)
            {
                var h = headers[c].Trim().TrimStart('\'').ToLower();
                foreach (var kw in codeKeywords)
                    if (h.Contains(kw)) return c;
            }

            // 2차: 데이터 패턴으로 찾기 (6자리 숫자가 가장 많은 컬럼)
            if (lines.Length < 2) return -1;

            var sampleLines = lines.Length > 11 ? 10 : lines.Length - 1;
            var counts      = new int[headers.Length];

            for (int i = 1; i <= sampleLines; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var cols = SplitCsv(lines[i]);
                for (int c = 0; c < Math.Min(cols.Length, headers.Length); c++)
                {
                    var val = cols[c].Trim().TrimStart('\'');
                    if (IsValidCode(val.PadLeft(6, '0'))) counts[c]++;
                }
            }

            int best = -1, bestCount = 0;
            for (int c = 0; c < counts.Length; c++)
                if (counts[c] > bestCount) { bestCount = counts[c]; best = c; }

            return bestCount >= sampleLines / 2 ? best : -1;
        }

        private static bool IsValidCode(string s)
            => s.Length == 6 && Regex.IsMatch(s, @"^\d{6}$");

        private static Encoding DetectEncoding(string path)
        {
            try
            {
                var bom = new byte[3];
                using (var fs = File.OpenRead(path))
                    fs.Read(bom, 0, 3);
                if (bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
                    return Encoding.UTF8;
            }
            catch { }
            return Encoding.GetEncoding(949);  // CP949 (EUC-KR)
        }

        private static string[] SplitCsv(string line)
        {
            var result  = new List<string>();
            bool inQ    = false;
            var cur     = new StringBuilder();
            foreach (char c in line)
            {
                if      (c == '"')       inQ = !inQ;
                else if (c == ',' && !inQ) { result.Add(cur.ToString()); cur.Clear(); }
                else                       cur.Append(c);
            }
            result.Add(cur.ToString());
            return result.ToArray();
        }
    }
}
