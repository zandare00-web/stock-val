using System;
using System.IO;
using Newtonsoft.Json;

namespace StockAnalyzer.Models
{
    public class ScoreConfig
    {
        // ── 기업가치 ─────────────────────────────────────────────
        public double PerScore  { get; set; } = 10.0;  // PER 할인율 만점
        public double PbrScore  { get; set; } = 8.0;   // PBR 할인율 만점
        public double RoeScore  { get; set; } = 7.0;   // ROE 만점

        // ── 종목수급 ─────────────────────────────────────────────
        public double ForeignD1Score  { get; set; } = 10.0;  // 외국인 당일
        public double Foreign5DScore  { get; set; } = 10.0;  // 외국인 5일
        public double Foreign20DScore { get; set; } = 10.0;  // 외국인 20일
        public double InstD1Score     { get; set; } = 10.0;  // 기관 당일
        public double Inst5DScore     { get; set; } = 10.0;  // 기관 5일
        public double Inst20DScore    { get; set; } = 10.0;  // 기관 20일
        public double TurnoverScore   { get; set; } = 10.0;  // 거래회전율 추세

        // ── 업종수급 ─────────────────────────────────────────────
        public double SectorForeignD1Score  { get; set; } = 5.0;
        public double SectorForeign5DScore  { get; set; } = 5.0;
        public double SectorForeign20DScore { get; set; } = 5.0;
        public double SectorInstD1Score     { get; set; } = 5.0;
        public double SectorInst5DScore     { get; set; } = 5.0;
        public double SectorInst20DScore    { get; set; } = 5.0;

        // ── 수급강도 추세 기준 ───────────────────────────────────
        public double TrendThresholdPct { get; set; } = 10.0;  // 보합 기준 (%)

        // ── 거래회전율 만점 기준 ─────────────────────────────────
        public double TurnoverFullPct { get; set; } = 50.0;  // 20일이 60일 대비 50% 이상 증가시 만점

        // ── KRX Open API 인증키 ─────────────────────────────────
        public string KrxAuthKey { get; set; } = "76180EC87B174F449DC331AE0B81158861A1136B";

        // ── 집계 ─────────────────────────────────────────────────
        [JsonIgnore]
        public double TotalValueScore =>
            PerScore + PbrScore + RoeScore;

        [JsonIgnore]
        public double TotalStockSupplyScore =>
            ForeignD1Score + Foreign5DScore + Foreign20DScore +
            InstD1Score + Inst5DScore + Inst20DScore + TurnoverScore;

        [JsonIgnore]
        public double TotalSectorSupplyScore =>
            SectorForeignD1Score + SectorForeign5DScore + SectorForeign20DScore +
            SectorInstD1Score + SectorInst5DScore + SectorInst20DScore;

        [JsonIgnore]
        public double TotalMaxScore =>
            TotalValueScore + TotalStockSupplyScore + TotalSectorSupplyScore;

        // ── 저장/불러오기 ────────────────────────────────────────
        private static readonly string _path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "StockAnalyzer", "score_config.json");

        private static ScoreConfig _instance;
        public static ScoreConfig Instance
        {
            get
            {
                if (_instance == null) _instance = Load();
                return _instance;
            }
        }

        public static ScoreConfig Load()
        {
            try
            {
                if (File.Exists(_path))
                    return JsonConvert.DeserializeObject<ScoreConfig>(File.ReadAllText(_path))
                           ?? new ScoreConfig();
            }
            catch { }
            return new ScoreConfig();
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_path));
                File.WriteAllText(_path, JsonConvert.SerializeObject(this, Formatting.Indented));
                _instance = this;
            }
            catch { }
        }
    }
}
