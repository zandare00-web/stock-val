using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq;
using StockAnalyzer.Models;

namespace StockAnalyzer.Cache
{
    /// <summary>
    /// 정확도 우선 로컬 캐시 DB (SQLite)
    /// - 일별 확정값 + 장중 스냅샷 + 수집 로그 저장
    /// - 분석 계산은 가능하면 DB 캐시 기준으로 수행
    /// </summary>
    public sealed class MarketCacheDb : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connStr;
        private bool _initialized;

        public string DbPath => _dbPath;

        public MarketCacheDb(string dbPath = null)
        {
            _dbPath = dbPath ?? DefaultPath();
            _connStr = $"Data Source={_dbPath};Version=3;Pooling=True;Journal Mode=WAL;Synchronous=Normal;";
        }

        public static string DefaultPath()
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "StockAnalyzer", "cache");
            return Path.Combine(dir, "market_cache_v1.sqlite");
        }

        public void EnsureCreated()
        {
            if (_initialized) return;
            Directory.CreateDirectory(Path.GetDirectoryName(_dbPath));
            if (!File.Exists(_dbPath)) SQLiteConnection.CreateFile(_dbPath);
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
PRAGMA journal_mode=WAL;
PRAGMA synchronous=NORMAL;
PRAGMA foreign_keys=OFF;

CREATE TABLE IF NOT EXISTS stock_master (
    code TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    market TEXT NOT NULL,
    krx_sector_name TEXT,
    kiwoom_sector_name TEXT,
    kiwoom_sector_code TEXT,
    current_price REAL,
    updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_stock_master_market ON stock_master(market);

CREATE TABLE IF NOT EXISTS stock_sector_map_history (
    code TEXT NOT NULL,
    trade_date TEXT NOT NULL,
    market TEXT NOT NULL,
    sector_code TEXT,
    sector_name TEXT,
    source TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (code, trade_date)
);

CREATE TABLE IF NOT EXISTS stock_investor_daily (
    code TEXT NOT NULL,
    trade_date TEXT NOT NULL,
    foreign_net_qty INTEGER,
    inst_net_qty INTEGER,
    foreign_net_amt INTEGER,
    inst_net_amt INTEGER,
    total_volume INTEGER,
    total_trading_value INTEGER,
    qty_status TEXT NOT NULL DEFAULT 'MISSING',
    amt_status TEXT NOT NULL DEFAULT 'MISSING',
    pair_status TEXT NOT NULL DEFAULT 'PARTIAL',
    qty_collected_at TEXT,
    amt_collected_at TEXT,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (code, trade_date)
);
CREATE INDEX IF NOT EXISTS idx_stock_inv_daily_date ON stock_investor_daily(trade_date);
CREATE INDEX IF NOT EXISTS idx_stock_inv_daily_code_date ON stock_investor_daily(code, trade_date);

CREATE TABLE IF NOT EXISTS sector_investor_daily (
    market TEXT NOT NULL,
    sector_code TEXT NOT NULL,
    sector_name TEXT,
    trade_date TEXT NOT NULL,
    foreign_net_amt INTEGER,
    inst_net_amt INTEGER,
    status TEXT NOT NULL DEFAULT 'MISSING',
    collected_at TEXT,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (market, sector_code, trade_date)
);
CREATE INDEX IF NOT EXISTS idx_sector_inv_daily_date ON sector_investor_daily(trade_date);

CREATE TABLE IF NOT EXISTS stock_daily_bar (
    code TEXT NOT NULL,
    trade_date TEXT NOT NULL,
    open INTEGER,
    high INTEGER,
    low INTEGER,
    close INTEGER,
    volume INTEGER,
    trading_value INTEGER,
    market_cap INTEGER,
    status TEXT NOT NULL DEFAULT 'OK',
    collected_at TEXT,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (code, trade_date)
);
CREATE INDEX IF NOT EXISTS idx_stock_daily_bar_code_date ON stock_daily_bar(code, trade_date);

CREATE TABLE IF NOT EXISTS stock_fundamental_snapshot (
    code TEXT NOT NULL,
    asof_date TEXT NOT NULL,
    per REAL,
    pbr REAL,
    roe REAL,
    eps REAL,
    bps REAL,
    market_cap REAL,
    status TEXT NOT NULL DEFAULT 'OK',
    source TEXT NOT NULL DEFAULT 'kiwoom',
    collected_at TEXT,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (code, asof_date)
);

CREATE TABLE IF NOT EXISTS sector_valuation_daily (
    market TEXT NOT NULL,
    sector_code TEXT NOT NULL,
    trade_date TEXT NOT NULL,
    per REAL,
    pbr REAL,
    status TEXT NOT NULL DEFAULT 'OK',
    collected_at TEXT,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (market, sector_code, trade_date)
);

CREATE TABLE IF NOT EXISTS realtime_quote_snapshot (
    code TEXT PRIMARY KEY,
    trade_date TEXT NOT NULL,
    current_price INTEGER,
    change_price INTEGER,
    change_rate REAL,
    cum_volume INTEGER,
    cum_trading_value INTEGER,
    high INTEGER,
    low INTEGER,
    last_tick_time TEXT,
    source TEXT NOT NULL DEFAULT 'polling',
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS realtime_investor_intraday (
    code TEXT NOT NULL,
    trade_date TEXT NOT NULL,
    asof_time TEXT NOT NULL,
    foreign_net_qty INTEGER,
    inst_net_qty INTEGER,
    foreign_net_amt INTEGER,
    inst_net_amt INTEGER,
    qty_status TEXT NOT NULL,
    amt_status TEXT NOT NULL,
    source TEXT NOT NULL DEFAULT 'opt10059',
    collected_at TEXT NOT NULL,
    PRIMARY KEY (code, trade_date, asof_time)
);
CREATE INDEX IF NOT EXISTS idx_rt_inv_code_date ON realtime_investor_intraday(code, trade_date, asof_time DESC);

CREATE TABLE IF NOT EXISTS realtime_sector_intraday (
    market TEXT NOT NULL,
    sector_code TEXT NOT NULL,
    trade_date TEXT NOT NULL,
    asof_time TEXT NOT NULL,
    foreign_net_amt INTEGER,
    inst_net_amt INTEGER,
    status TEXT NOT NULL,
    source TEXT NOT NULL DEFAULT 'opt10051',
    collected_at TEXT NOT NULL,
    PRIMARY KEY (market, sector_code, trade_date, asof_time)
);
CREATE INDEX IF NOT EXISTS idx_rt_sector_date ON realtime_sector_intraday(market, trade_date, asof_time DESC);

CREATE TABLE IF NOT EXISTS fetch_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at TEXT NOT NULL,
    ended_at TEXT,
    source TEXT NOT NULL,
    tr_code TEXT,
    rq_name TEXT,
    target_key TEXT,
    params_json TEXT,
    status TEXT NOT NULL,
    record_count INTEGER,
    error_message TEXT
);
CREATE INDEX IF NOT EXISTS idx_fetch_log_started ON fetch_log(started_at);
CREATE INDEX IF NOT EXISTS idx_fetch_log_tr ON fetch_log(tr_code, started_at);
";
                cmd.ExecuteNonQuery();
            }
            _initialized = true;
        }

        private SQLiteConnection Open()
        {
            var con = new SQLiteConnection(_connStr);
            con.Open();
            return con;
        }

        public void Dispose() { }

        private static string NowIso() => DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        private static string Ymd(DateTime d) => d.ToString("yyyyMMdd");

        public void UpsertStockInfos(IEnumerable<StockInfo> stocks)
        {
            EnsureCreated();
            if (stocks == null) return;
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                foreach (var s in stocks)
                {
                    if (s == null || string.IsNullOrWhiteSpace(s.Code)) continue;
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
INSERT INTO stock_master(code,name,market,krx_sector_name,current_price,updated_at)
VALUES(@code,@name,@market,@sec,@cp,@u)
ON CONFLICT(code) DO UPDATE SET
 name=excluded.name,
 market=excluded.market,
 krx_sector_name=COALESCE(NULLIF(excluded.krx_sector_name,''), stock_master.krx_sector_name),
 current_price=CASE WHEN excluded.current_price IS NULL OR excluded.current_price<=0 THEN stock_master.current_price ELSE excluded.current_price END,
 updated_at=excluded.updated_at;";
                        cmd.Parameters.AddWithValue("@code", s.Code);
                        cmd.Parameters.AddWithValue("@name", s.Name ?? s.Code);
                        cmd.Parameters.AddWithValue("@market", string.IsNullOrWhiteSpace(s.Market) ? "KOSPI" : s.Market);
                        cmd.Parameters.AddWithValue("@sec", (object)(s.SectorName ?? ""));
                        cmd.Parameters.AddWithValue("@cp", s.CurrentPrice);
                        cmd.Parameters.AddWithValue("@u", NowIso());
                        cmd.ExecuteNonQuery();
                    }
                    if (s.CurrentPrice > 0)
                    {
                        UpsertRealtimeQuoteSnapshot(new RealtimeQuoteSnapshotRow
                        {
                            Code = s.Code,
                            TradeDate = Ymd(DateTime.Today),
                            CurrentPrice = (long)Math.Round(s.CurrentPrice),
                            CumVolume = null,
                            CumTradingValue = null,
                            LastTickTime = DateTime.Now.ToString("HHmmss"),
                            Source = "krx_polling"
                        }, con, tx);
                    }
                }
                tx.Commit();
            }
        }

        public void UpsertStockSectorMap(string code, DateTime tradeDate, string market, string sectorCode, string sectorName, string source = "KOA_Functions")
        {
            EnsureCreated();
            if (string.IsNullOrWhiteSpace(code)) return;
            var ymd = Ymd(tradeDate);
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"
INSERT INTO stock_sector_map_history(code,trade_date,market,sector_code,sector_name,source,updated_at)
VALUES(@c,@d,@m,@sc,@sn,@src,@u)
ON CONFLICT(code,trade_date) DO UPDATE SET
 market=excluded.market, sector_code=excluded.sector_code, sector_name=excluded.sector_name, source=excluded.source, updated_at=excluded.updated_at;";
                    cmd.Parameters.AddWithValue("@c", code);
                    cmd.Parameters.AddWithValue("@d", ymd);
                    cmd.Parameters.AddWithValue("@m", string.IsNullOrWhiteSpace(market) ? "KOSPI" : market);
                    cmd.Parameters.AddWithValue("@sc", (object)(sectorCode ?? ""));
                    cmd.Parameters.AddWithValue("@sn", (object)(sectorName ?? ""));
                    cmd.Parameters.AddWithValue("@src", source ?? "KOA_Functions");
                    cmd.Parameters.AddWithValue("@u", NowIso());
                    cmd.ExecuteNonQuery();
                }
                using (var cmd2 = con.CreateCommand())
                {
                    cmd2.Transaction = tx;
                    cmd2.CommandText = @"
UPDATE stock_master
SET market=COALESCE(NULLIF(@m,''), market),
    kiwoom_sector_name=COALESCE(NULLIF(@sn,''), kiwoom_sector_name),
    kiwoom_sector_code=COALESCE(NULLIF(@sc,''), kiwoom_sector_code),
    updated_at=@u
WHERE code=@c;";
                    cmd2.Parameters.AddWithValue("@c", code);
                    cmd2.Parameters.AddWithValue("@m", market ?? "");
                    cmd2.Parameters.AddWithValue("@sn", sectorName ?? "");
                    cmd2.Parameters.AddWithValue("@sc", sectorCode ?? "");
                    cmd2.Parameters.AddWithValue("@u", NowIso());
                    cmd2.ExecuteNonQuery();
                }
                tx.Commit();
            }
        }

        public FundamentalData GetLatestFundamental(string code)
        {
            EnsureCreated();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"SELECT code,per,pbr,roe,eps,bps,market_cap FROM stock_fundamental_snapshot WHERE code=@c AND status='OK' AND asof_date>=@minD ORDER BY asof_date DESC LIMIT 1;";
                cmd.Parameters.AddWithValue("@minD", DateTime.Today.AddDays(-7).ToString("yyyyMMdd"));
                cmd.Parameters.AddWithValue("@c", code);
                using (var rd = cmd.ExecuteReader())
                {
                    if (!rd.Read()) return null;
                    return new FundamentalData
                    {
                        Code = code,
                        Per = DbNullD(rd, 1),
                        Pbr = DbNullD(rd, 2),
                        Roe = DbNullD(rd, 3),
                        Eps = DbNullD(rd, 4),
                        Bps = DbNullD(rd, 5),
                        MarketCap = DbD(rd, 6)
                    };
                }
            }
        }

        public void UpsertFundamental(string code, FundamentalData f, string asofDateYmd, string status = "OK", string source = "kiwoom")
        {
            EnsureCreated();
            if (f == null || string.IsNullOrWhiteSpace(code)) return;
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
INSERT INTO stock_fundamental_snapshot(code,asof_date,per,pbr,roe,eps,bps,market_cap,status,source,collected_at,updated_at)
VALUES(@c,@d,@per,@pbr,@roe,@eps,@bps,@mc,@st,@src,@t,@u)
ON CONFLICT(code,asof_date) DO UPDATE SET
 per=excluded.per,pbr=excluded.pbr,roe=excluded.roe,eps=excluded.eps,bps=excluded.bps,market_cap=excluded.market_cap,
 status=excluded.status,source=excluded.source,collected_at=excluded.collected_at,updated_at=excluded.updated_at;";
                cmd.Parameters.AddWithValue("@c", code);
                cmd.Parameters.AddWithValue("@d", asofDateYmd);
                AddNullable(cmd, "@per", f.Per);
                AddNullable(cmd, "@pbr", f.Pbr);
                AddNullable(cmd, "@roe", f.Roe);
                AddNullable(cmd, "@eps", f.Eps);
                AddNullable(cmd, "@bps", f.Bps);
                cmd.Parameters.AddWithValue("@mc", f.MarketCap);
                cmd.Parameters.AddWithValue("@st", status ?? "OK");
                cmd.Parameters.AddWithValue("@src", source ?? "kiwoom");
                cmd.Parameters.AddWithValue("@t", NowIso());
                cmd.Parameters.AddWithValue("@u", NowIso());
                cmd.ExecuteNonQuery();
            }
        }

        public List<InvestorDay> GetStockInvestorDaily(string code, DateTime fromDate)
        {
            EnsureCreated();
            var list = new List<InvestorDay>();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT trade_date, foreign_net_qty, inst_net_qty, foreign_net_amt, inst_net_amt, total_volume, total_trading_value
FROM stock_investor_daily
WHERE code=@c AND trade_date>=@d AND pair_status IN ('COMPLETE','PARTIAL')
ORDER BY trade_date DESC;";
                cmd.Parameters.AddWithValue("@c", code);
                cmd.Parameters.AddWithValue("@d", Ymd(fromDate));
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        if (!DateTime.TryParseExact((rd[0]?.ToString() ?? "").Trim(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                            continue;
                        var row = new InvestorDay
                        {
                            Date = dt,
                            ForeignNetQty = DbL(rd, 1),
                            InstNetQty = DbL(rd, 2),
                            ForeignNetAmt = DbL(rd, 3),
                            InstNetAmt = DbL(rd, 4),
                            Volume = DbL(rd, 5),
                            TradeAmount = DbD(rd, 6)
                        };
                        // 종목수급은 수량 기준으로 사용 (금액 필드는 비활성)
                        row.ForeignNet = row.ForeignNetQty != 0 ? row.ForeignNetQty : row.ForeignNetAmt;
                        row.InstNet = row.InstNetQty != 0 ? row.InstNetQty : row.InstNetAmt;
                        list.Add(row);
                    }
                }
            }
            return list;
        }

        public void UpsertStockInvestorDaily(string code, IEnumerable<InvestorDay> rows)
        {
            EnsureCreated();
            if (string.IsNullOrWhiteSpace(code) || rows == null) return;
            var now = NowIso();
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                foreach (var x in rows)
                {
                    if (x == null) continue;
                    var ymd = Ymd(x.Date);
                    var hasAmt = x.ForeignNetAmt != 0 || x.InstNetAmt != 0;
                    var hasQty = x.ForeignNetQty != 0 || x.InstNetQty != 0;
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
INSERT INTO stock_investor_daily(code,trade_date,foreign_net_qty,inst_net_qty,foreign_net_amt,inst_net_amt,total_volume,total_trading_value,qty_status,amt_status,pair_status,qty_collected_at,amt_collected_at,updated_at)
VALUES(@c,@d,@fq,@iq,@fa,@ia,@vol,@tv,@qs,@as,@ps,@qt,@at,@u)
ON CONFLICT(code,trade_date) DO UPDATE SET
 foreign_net_qty=COALESCE(excluded.foreign_net_qty, stock_investor_daily.foreign_net_qty),
 inst_net_qty=COALESCE(excluded.inst_net_qty, stock_investor_daily.inst_net_qty),
 foreign_net_amt=COALESCE(excluded.foreign_net_amt, stock_investor_daily.foreign_net_amt),
 inst_net_amt=COALESCE(excluded.inst_net_amt, stock_investor_daily.inst_net_amt),
 total_volume=CASE WHEN excluded.total_volume IS NULL OR excluded.total_volume=0 THEN stock_investor_daily.total_volume ELSE excluded.total_volume END,
 total_trading_value=CASE WHEN excluded.total_trading_value IS NULL OR excluded.total_trading_value=0 THEN stock_investor_daily.total_trading_value ELSE excluded.total_trading_value END,
 qty_status=CASE WHEN excluded.qty_status='OK' THEN 'OK' ELSE stock_investor_daily.qty_status END,
 amt_status=CASE WHEN excluded.amt_status='OK' THEN 'OK' ELSE stock_investor_daily.amt_status END,
 pair_status=CASE
   WHEN (CASE WHEN excluded.qty_status='OK' THEN 'OK' ELSE stock_investor_daily.qty_status END)='OK' AND (CASE WHEN excluded.amt_status='OK' THEN 'OK' ELSE stock_investor_daily.amt_status END)='OK' THEN 'COMPLETE'
   WHEN (CASE WHEN excluded.qty_status='OK' THEN 'OK' ELSE stock_investor_daily.qty_status END)='OK' OR (CASE WHEN excluded.amt_status='OK' THEN 'OK' ELSE stock_investor_daily.amt_status END)='OK' THEN 'PARTIAL'
   ELSE stock_investor_daily.pair_status END,
 qty_collected_at=COALESCE(excluded.qty_collected_at, stock_investor_daily.qty_collected_at),
 amt_collected_at=COALESCE(excluded.amt_collected_at, stock_investor_daily.amt_collected_at),
 updated_at=excluded.updated_at;";
                        cmd.Parameters.AddWithValue("@c", code);
                        cmd.Parameters.AddWithValue("@d", ymd);
                        AddNullable(cmd, "@fq", x.ForeignNetQty == 0 ? (long?)null : x.ForeignNetQty);
                        AddNullable(cmd, "@iq", x.InstNetQty == 0 ? (long?)null : x.InstNetQty);
                        AddNullable(cmd, "@fa", x.ForeignNetAmt == 0 ? (long?)null : x.ForeignNetAmt);
                        AddNullable(cmd, "@ia", x.InstNetAmt == 0 ? (long?)null : x.InstNetAmt);
                        AddNullable(cmd, "@vol", x.Volume == 0 ? (long?)null : x.Volume);
                        AddNullable(cmd, "@tv", x.TradeAmount == 0 ? (long?)null : (long)Math.Round(x.TradeAmount));
                        cmd.Parameters.AddWithValue("@qs", hasQty ? "OK" : "MISSING");
                        cmd.Parameters.AddWithValue("@as", hasAmt ? "OK" : "MISSING");
                        cmd.Parameters.AddWithValue("@ps", hasQty && hasAmt ? "COMPLETE" : hasQty || hasAmt ? "PARTIAL" : "MISSING");
                        AddNullable(cmd, "@qt", hasQty ? now : null);
                        AddNullable(cmd, "@at", hasAmt ? now : null);
                        cmd.Parameters.AddWithValue("@u", now);
                        cmd.ExecuteNonQuery();
                    }
                }
                tx.Commit();
            }
        }

        public void UpsertRealtimeInvestorSnapshot(string code, InvestorDay todayRow, DateTime asof)
        {
            EnsureCreated();
            if (todayRow == null || string.IsNullOrWhiteSpace(code)) return;
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
INSERT OR REPLACE INTO realtime_investor_intraday(code,trade_date,asof_time,foreign_net_qty,inst_net_qty,foreign_net_amt,inst_net_amt,qty_status,amt_status,source,collected_at)
VALUES(@c,@d,@t,@fq,@iq,@fa,@ia,@qs,@as,'opt10059',@col);";
                cmd.Parameters.AddWithValue("@c", code);
                cmd.Parameters.AddWithValue("@d", Ymd(todayRow.Date));
                cmd.Parameters.AddWithValue("@t", asof.ToString("HHmmss"));
                AddNullable(cmd, "@fq", todayRow.ForeignNetQty == 0 ? (long?)null : todayRow.ForeignNetQty);
                AddNullable(cmd, "@iq", todayRow.InstNetQty == 0 ? (long?)null : todayRow.InstNetQty);
                AddNullable(cmd, "@fa", todayRow.ForeignNetAmt == 0 ? (long?)null : todayRow.ForeignNetAmt);
                AddNullable(cmd, "@ia", todayRow.InstNetAmt == 0 ? (long?)null : todayRow.InstNetAmt);
                cmd.Parameters.AddWithValue("@qs", (todayRow.ForeignNetQty != 0 || todayRow.InstNetQty != 0) ? "OK" : "MISSING");
                cmd.Parameters.AddWithValue("@as", (todayRow.ForeignNetAmt != 0 || todayRow.InstNetAmt != 0) ? "OK" : "MISSING");
                cmd.Parameters.AddWithValue("@col", NowIso());
                cmd.ExecuteNonQuery();
            }
        }

        public InvestorDay GetLatestRealtimeInvestorSnapshot(string code, DateTime tradeDate)
        {
            EnsureCreated();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT foreign_net_qty,inst_net_qty,foreign_net_amt,inst_net_amt
FROM realtime_investor_intraday
WHERE code=@c AND trade_date=@d
ORDER BY asof_time DESC LIMIT 1;";
                cmd.Parameters.AddWithValue("@c", code);
                cmd.Parameters.AddWithValue("@d", Ymd(tradeDate));
                using (var rd = cmd.ExecuteReader())
                {
                    if (!rd.Read()) return null;
                    return new InvestorDay
                    {
                        Date = tradeDate.Date,
                        ForeignNetQty = DbL(rd, 0),
                        InstNetQty = DbL(rd, 1),
                        ForeignNetAmt = DbL(rd, 2),
                        InstNetAmt = DbL(rd, 3),
                        ForeignNet = DbL(rd, 0) != 0 ? DbL(rd, 0) : DbL(rd, 2),
                        InstNet = DbL(rd, 1) != 0 ? DbL(rd, 1) : DbL(rd, 3),
                    };
                }
            }
        }

        public List<DailyBar> GetDailyBars(string code, DateTime fromDate)
        {
            EnsureCreated();
            var list = new List<DailyBar>();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT trade_date, market_cap, trading_value, volume
FROM stock_daily_bar
WHERE code=@c AND trade_date>=@d AND status='OK'
ORDER BY trade_date DESC;";
                cmd.Parameters.AddWithValue("@c", code);
                cmd.Parameters.AddWithValue("@d", Ymd(fromDate));
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        if (!DateTime.TryParseExact((rd[0]?.ToString() ?? "").Trim(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                            continue;
                        list.Add(new DailyBar
                        {
                            Date = dt,
                            MarketCap = DbD(rd, 1),
                            TradeAmount = DbD(rd, 2),
                            Volume = DbL(rd, 3)
                        });
                    }
                }
            }
            return list;
        }

        public void UpsertDailyBars(string code, IEnumerable<DailyBar> rows)
        {
            EnsureCreated();
            if (string.IsNullOrWhiteSpace(code) || rows == null) return;
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                foreach (var b in rows)
                {
                    if (b == null) continue;
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
INSERT INTO stock_daily_bar(code,trade_date,volume,trading_value,market_cap,status,collected_at,updated_at)
VALUES(@c,@d,@v,@tv,@mc,'OK',@t,@u)
ON CONFLICT(code,trade_date) DO UPDATE SET
 volume=excluded.volume,
 trading_value=excluded.trading_value,
 market_cap=CASE WHEN excluded.market_cap IS NULL OR excluded.market_cap<=0 THEN stock_daily_bar.market_cap ELSE excluded.market_cap END,
 status='OK', collected_at=excluded.collected_at, updated_at=excluded.updated_at;";
                        cmd.Parameters.AddWithValue("@c", code);
                        cmd.Parameters.AddWithValue("@d", Ymd(b.Date));
                        AddNullable(cmd, "@v", b.Volume == 0 ? (long?)null : b.Volume);
                        AddNullable(cmd, "@tv", b.TradeAmount == 0 ? (long?)null : (long)Math.Round(b.TradeAmount));
                        AddNullable(cmd, "@mc", b.MarketCap == 0 ? (long?)null : (long)Math.Round(b.MarketCap));
                        cmd.Parameters.AddWithValue("@t", NowIso());
                        cmd.Parameters.AddWithValue("@u", NowIso());
                        cmd.ExecuteNonQuery();
                    }
                }
                tx.Commit();
            }
        }

        public List<SectorInvestorRow> GetSectorInvestorDaily(string market, DateTime tradeDate)
        {
            EnsureCreated();
            var list = new List<SectorInvestorRow>();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT sector_code, sector_name, foreign_net_amt, inst_net_amt
FROM sector_investor_daily
WHERE market=@m AND trade_date=@d AND status='OK'
ORDER BY sector_code;";
                cmd.Parameters.AddWithValue("@m", market ?? "KOSPI");
                cmd.Parameters.AddWithValue("@d", Ymd(tradeDate));
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        list.Add(new SectorInvestorRow
                        {
                            SectorCode = (rd[0]?.ToString() ?? "").Trim(),
                            SectorName = (rd[1]?.ToString() ?? "").Trim(),
                            ForeignNet = DbD(rd, 2),
                            InstNet = DbD(rd, 3),
                        });
                    }
                }
            }
            return list;
        }

        public void UpsertSectorInvestorDaily(string market, DateTime tradeDate, IEnumerable<SectorInvestorRow> rows, string status = "OK")
        {
            EnsureCreated();
            if (rows == null) return;
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                foreach (var r in rows)
                {
                    if (r == null || string.IsNullOrWhiteSpace(r.SectorCode)) continue;
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
INSERT INTO sector_investor_daily(market,sector_code,sector_name,trade_date,foreign_net_amt,inst_net_amt,status,collected_at,updated_at)
VALUES(@m,@sc,@sn,@d,@f,@i,@st,@t,@u)
ON CONFLICT(market,sector_code,trade_date) DO UPDATE SET
 sector_name=COALESCE(NULLIF(excluded.sector_name,''), sector_investor_daily.sector_name),
 foreign_net_amt=excluded.foreign_net_amt,
 inst_net_amt=excluded.inst_net_amt,
 status=excluded.status,
 collected_at=excluded.collected_at,
 updated_at=excluded.updated_at;";
                        cmd.Parameters.AddWithValue("@m", market ?? "KOSPI");
                        cmd.Parameters.AddWithValue("@sc", r.SectorCode?.Trim() ?? "");
                        cmd.Parameters.AddWithValue("@sn", r.SectorName ?? "");
                        cmd.Parameters.AddWithValue("@d", Ymd(tradeDate));
                        cmd.Parameters.AddWithValue("@f", (long)Math.Round(r.ForeignNet));
                        cmd.Parameters.AddWithValue("@i", (long)Math.Round(r.InstNet));
                        cmd.Parameters.AddWithValue("@st", status ?? "OK");
                        cmd.Parameters.AddWithValue("@t", NowIso());
                        cmd.Parameters.AddWithValue("@u", NowIso());
                        cmd.ExecuteNonQuery();
                    }
                }
                tx.Commit();
            }
        }

        public void UpsertRealtimeSectorSnapshot(string market, DateTime tradeDate, IEnumerable<SectorInvestorRow> rows, DateTime asof)
        {
            EnsureCreated();
            if (rows == null) return;
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                foreach (var r in rows)
                {
                    if (r == null || string.IsNullOrWhiteSpace(r.SectorCode)) continue;
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
INSERT OR REPLACE INTO realtime_sector_intraday(market,sector_code,trade_date,asof_time,foreign_net_amt,inst_net_amt,status,source,collected_at)
VALUES(@m,@sc,@d,@t,@f,@i,'OK','opt10051',@col);";
                        cmd.Parameters.AddWithValue("@m", market ?? "KOSPI");
                        cmd.Parameters.AddWithValue("@sc", r.SectorCode?.Trim() ?? "");
                        cmd.Parameters.AddWithValue("@d", Ymd(tradeDate));
                        cmd.Parameters.AddWithValue("@t", asof.ToString("HHmmss"));
                        cmd.Parameters.AddWithValue("@f", (long)Math.Round(r.ForeignNet));
                        cmd.Parameters.AddWithValue("@i", (long)Math.Round(r.InstNet));
                        cmd.Parameters.AddWithValue("@col", NowIso());
                        cmd.ExecuteNonQuery();
                    }
                }
                tx.Commit();
            }
        }

        public Dictionary<string, SectorInvestorRow> GetLatestRealtimeSectorSnapshot(string market, DateTime tradeDate)
        {
            EnsureCreated();
            var map = new Dictionary<string, SectorInvestorRow>();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT t1.sector_code, t1.foreign_net_amt, t1.inst_net_amt
FROM realtime_sector_intraday t1
INNER JOIN (
  SELECT sector_code, MAX(asof_time) AS max_t
  FROM realtime_sector_intraday
  WHERE market=@m AND trade_date=@d
  GROUP BY sector_code
) t2 ON t1.sector_code=t2.sector_code AND t1.asof_time=t2.max_t
WHERE t1.market=@m AND t1.trade_date=@d;";
                cmd.Parameters.AddWithValue("@m", market ?? "KOSPI");
                cmd.Parameters.AddWithValue("@d", Ymd(tradeDate));
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        var sc = (rd[0]?.ToString() ?? "").Trim();
                        if (string.IsNullOrEmpty(sc)) continue;
                        map[sc] = new SectorInvestorRow
                        {
                            SectorCode = sc,
                            ForeignNet = DbD(rd, 1),
                            InstNet = DbD(rd, 2)
                        };
                    }
                }
            }
            return map;
        }

        public SectorFundamental GetSectorValuation(string market, string sectorCode, DateTime tradeDate)
        {
            EnsureCreated();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT per,pbr FROM sector_valuation_daily
WHERE market=@m AND sector_code=@s AND trade_date=@d AND status='OK'
LIMIT 1;";
                cmd.Parameters.AddWithValue("@m", market ?? "KOSPI");
                cmd.Parameters.AddWithValue("@s", sectorCode ?? "");
                cmd.Parameters.AddWithValue("@d", Ymd(tradeDate));
                using (var rd = cmd.ExecuteReader())
                {
                    if (!rd.Read()) return null;
                    return new SectorFundamental { SectorCode = sectorCode, AvgPer = DbNullD(rd, 0), AvgPbr = DbNullD(rd, 1) };
                }
            }
        }

        public void UpsertSectorValuation(string market, string sectorCode, string sectorName, DateTime tradeDate, SectorFundamental sf, string status = "OK")
        {
            EnsureCreated();
            if (string.IsNullOrWhiteSpace(sectorCode) || sf == null) return;
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
INSERT INTO sector_valuation_daily(market,sector_code,trade_date,per,pbr,status,collected_at,updated_at)
VALUES(@m,@s,@d,@per,@pbr,@st,@t,@u)
ON CONFLICT(market,sector_code,trade_date) DO UPDATE SET per=excluded.per,pbr=excluded.pbr,status=excluded.status,collected_at=excluded.collected_at,updated_at=excluded.updated_at;";
                cmd.Parameters.AddWithValue("@m", market ?? "KOSPI");
                cmd.Parameters.AddWithValue("@s", sectorCode);
                cmd.Parameters.AddWithValue("@d", Ymd(tradeDate));
                AddNullable(cmd, "@per", sf.AvgPer);
                AddNullable(cmd, "@pbr", sf.AvgPbr);
                cmd.Parameters.AddWithValue("@st", status ?? "OK");
                cmd.Parameters.AddWithValue("@t", NowIso());
                cmd.Parameters.AddWithValue("@u", NowIso());
                cmd.ExecuteNonQuery();
            }
        }

        public void UpsertRealtimeQuoteSnapshot(RealtimeQuoteSnapshotRow row)
        {
            EnsureCreated();
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                UpsertRealtimeQuoteSnapshot(row, con, tx);
                tx.Commit();
            }
        }

        private void UpsertRealtimeQuoteSnapshot(RealtimeQuoteSnapshotRow row, SQLiteConnection con, SQLiteTransaction tx)
        {
            if (row == null || string.IsNullOrWhiteSpace(row.Code)) return;
            using (var cmd = con.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = @"
INSERT INTO realtime_quote_snapshot(code,trade_date,current_price,change_price,change_rate,cum_volume,cum_trading_value,high,low,last_tick_time,source,updated_at)
VALUES(@c,@d,@p,@cp,@cr,@v,@tv,@h,@l,@lt,@src,@u)
ON CONFLICT(code) DO UPDATE SET
 trade_date=excluded.trade_date,
 current_price=COALESCE(excluded.current_price, realtime_quote_snapshot.current_price),
 change_price=COALESCE(excluded.change_price, realtime_quote_snapshot.change_price),
 change_rate=COALESCE(excluded.change_rate, realtime_quote_snapshot.change_rate),
 cum_volume=COALESCE(excluded.cum_volume, realtime_quote_snapshot.cum_volume),
 cum_trading_value=COALESCE(excluded.cum_trading_value, realtime_quote_snapshot.cum_trading_value),
 high=COALESCE(excluded.high, realtime_quote_snapshot.high),
 low=COALESCE(excluded.low, realtime_quote_snapshot.low),
 last_tick_time=COALESCE(excluded.last_tick_time, realtime_quote_snapshot.last_tick_time),
 source=excluded.source,
 updated_at=excluded.updated_at;";
                cmd.Parameters.AddWithValue("@c", row.Code);
                cmd.Parameters.AddWithValue("@d", row.TradeDate ?? Ymd(DateTime.Today));
                AddNullable(cmd, "@p", row.CurrentPrice);
                AddNullable(cmd, "@cp", row.ChangePrice);
                AddNullable(cmd, "@cr", row.ChangeRate);
                AddNullable(cmd, "@v", row.CumVolume);
                AddNullable(cmd, "@tv", row.CumTradingValue);
                AddNullable(cmd, "@h", row.High);
                AddNullable(cmd, "@l", row.Low);
                AddNullable(cmd, "@lt", row.LastTickTime);
                cmd.Parameters.AddWithValue("@src", row.Source ?? "polling");
                cmd.Parameters.AddWithValue("@u", NowIso());
                cmd.ExecuteNonQuery();
            }
        }

        public RealtimeQuoteSnapshotRow GetRealtimeQuoteSnapshot(string code)
        {
            EnsureCreated();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"SELECT trade_date,current_price,change_price,change_rate,cum_volume,cum_trading_value,high,low,last_tick_time,source FROM realtime_quote_snapshot WHERE code=@c LIMIT 1;";
                cmd.Parameters.AddWithValue("@c", code);
                using (var rd = cmd.ExecuteReader())
                {
                    if (!rd.Read()) return null;
                    return new RealtimeQuoteSnapshotRow
                    {
                        Code = code,
                        TradeDate = rd[0]?.ToString(),
                        CurrentPrice = DbNullL(rd, 1),
                        ChangePrice = DbNullL(rd, 2),
                        ChangeRate = DbNullD(rd, 3),
                        CumVolume = DbNullL(rd, 4),
                        CumTradingValue = DbNullL(rd, 5),
                        High = DbNullL(rd, 6),
                        Low = DbNullL(rd, 7),
                        LastTickTime = rd[8]?.ToString(),
                        Source = rd[9]?.ToString()
                    };
                }
            }
        }

        public void WriteFetchLog(string source, string trCode, string rqName, string targetKey, string paramsJson, string status, int recordCount = 0, string error = null, DateTime? startedAt = null, DateTime? endedAt = null)
        {
            EnsureCreated();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
INSERT INTO fetch_log(started_at,ended_at,source,tr_code,rq_name,target_key,params_json,status,record_count,error_message)
VALUES(@s,@e,@src,@tr,@rq,@tk,@pj,@st,@rc,@er);";
                cmd.Parameters.AddWithValue("@s", (startedAt ?? DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@e", (object)((endedAt ?? DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss")));
                cmd.Parameters.AddWithValue("@src", source ?? "");
                cmd.Parameters.AddWithValue("@tr", (object)(trCode ?? ""));
                cmd.Parameters.AddWithValue("@rq", (object)(rqName ?? ""));
                cmd.Parameters.AddWithValue("@tk", (object)(targetKey ?? ""));
                cmd.Parameters.AddWithValue("@pj", (object)(paramsJson ?? ""));
                cmd.Parameters.AddWithValue("@st", status ?? "OK");
                cmd.Parameters.AddWithValue("@rc", recordCount);
                cmd.Parameters.AddWithValue("@er", (object)(error ?? ""));
                cmd.ExecuteNonQuery();
            }
        }

        public string GetHealthSummary()
        {
            EnsureCreated();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT
 (SELECT COUNT(*) FROM stock_master),
 (SELECT COUNT(*) FROM stock_investor_daily),
 (SELECT COUNT(*) FROM stock_daily_bar),
 (SELECT COUNT(*) FROM sector_investor_daily),
 (SELECT COUNT(*) FROM realtime_investor_intraday),
 (SELECT COUNT(*) FROM realtime_sector_intraday);";
                using (var rd = cmd.ExecuteReader())
                {
                    if (!rd.Read()) return "cache n/a";
                    return $"cache={Path.GetFileName(_dbPath)} | 종목:{DbL(rd,0)} 종목수급:{DbL(rd,1)} 일봉:{DbL(rd,2)} 업종수급:{DbL(rd,3)} RT종목:{DbL(rd,4)} RT업종:{DbL(rd,5)}";
                }
            }
        }

        private static void AddNullable(SQLiteCommand cmd, string name, object value)
        {
            cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        }

        private static double DbD(IDataRecord rd, int idx) => rd.IsDBNull(idx) ? 0d : Convert.ToDouble(rd.GetValue(idx));
        private static double? DbNullD(IDataRecord rd, int idx) => rd.IsDBNull(idx) ? (double?)null : Convert.ToDouble(rd.GetValue(idx));
        private static long DbL(IDataRecord rd, int idx) => rd.IsDBNull(idx) ? 0L : Convert.ToInt64(rd.GetValue(idx));
        private static long? DbNullL(IDataRecord rd, int idx) => rd.IsDBNull(idx) ? (long?)null : Convert.ToInt64(rd.GetValue(idx));
    }

    public sealed class RealtimeQuoteSnapshotRow
    {
        public string Code { get; set; }
        public string TradeDate { get; set; }
        public long? CurrentPrice { get; set; }
        public long? ChangePrice { get; set; }
        public double? ChangeRate { get; set; }
        public long? CumVolume { get; set; }
        public long? CumTradingValue { get; set; }
        public long? High { get; set; }
        public long? Low { get; set; }
        public string LastTickTime { get; set; }
        public string Source { get; set; }
    }
}
