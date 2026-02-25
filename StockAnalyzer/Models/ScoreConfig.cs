using System;
using System.IO;
using Newtonsoft.Json;

namespace StockAnalyzer.Models
{
    public class ScoreConfig
    {
        // ── 기업가치 (총 5점) ────────────────────────────────────
        public double PerScore { get; set; } = 2.0;
        public double PbrScore { get; set; } = 1.5;
        public double RoeScore { get; set; } = 1.5;

        // ── 종목수급 (총 55점) ───────────────────────────────────
        // 외국인+기관 합계 기준
        public double Supply5DScore { get; set; } = 15.0;  // 합계 5일
        public double Supply10DScore { get; set; } = 15.0;  // 합계 10일
        public double Supply20DScore { get; set; } = 15.0;  // 합계 20일
        public double TurnoverScore { get; set; } = 10.0;  // 거래회전율 추세

        // ── 업종수급 (총 40점) ───────────────────────────────────
        // 외국인+기관 합계 기준
        public double SectorSupply5DScore { get; set; } = 20.0;  // 합계 5일
        public double SectorSupply10DScore { get; set; } = 20.0;  // 합계 10일

        // ── 기준값 ──────────────────────────────────────────────
        public double TrendThresholdPct { get; set; } = 10.0;
        public double TurnoverFullPct { get; set; } = 50.0;

        // ── KRX Open API ────────────────────────────────────────
        public string KrxAuthKey { get; set; } = "76180EC87B174F449DC331AE0B81158861A1136B";

        // ── 집계 ────────────────────────────────────────────────
        [JsonIgnore] public double TotalValueScore => PerScore + PbrScore + RoeScore;
        [JsonIgnore] public double TotalStockSupplyScore => Supply5DScore + Supply10DScore + Supply20DScore + TurnoverScore;
        [JsonIgnore] public double TotalSectorSupplyScore => SectorSupply5DScore + SectorSupply10DScore;
        [JsonIgnore] public double TotalMaxScore => TotalValueScore + TotalStockSupplyScore + TotalSectorSupplyScore;

        // ── 저장/불러오기 ───────────────────────────────────────
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