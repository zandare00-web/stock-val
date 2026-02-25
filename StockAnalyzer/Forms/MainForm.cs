using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using AxKHOpenAPILib;
using StockAnalyzer.Api;
using StockAnalyzer.Models;
using StockAnalyzer.Scoring;
using StockAnalyzer.Utils;

namespace StockAnalyzer.Forms
{
    // ══════════════════════════════════════════════════════════
    //  커스텀 컨트롤 
    // ══════════════════════════════════════════════════════════

    internal sealed class RndBtn : Control
    {
        public Color Bg, Fg, Bdr;
        public int Rad = 8;
        bool _h;
        public RndBtn(string t, Color bg, Color fg, int w, int h) { Text = t; Bg = bg; Fg = fg; Bdr = Color.Empty; Size = new Size(w, h); Font = new Font("Segoe UI", 9f, FontStyle.Bold); Cursor = Cursors.Hand; SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer | ControlStyles.SupportsTransparentBackColor, true); BackColor = Color.Transparent; }
        protected override void OnPaint(PaintEventArgs e) { var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias; var r = new Rectangle(0, 0, Width - 1, Height - 1); var bg = Enabled ? (_h ? Lt(Bg, -15) : Bg) : Color.FromArgb(226, 232, 240); using (var p = RR(r, Rad)) { using (var b = new SolidBrush(bg)) g.FillPath(b, p); if (Bdr != Color.Empty) using (var pen = new Pen(Bdr)) g.DrawPath(pen, p); } TextRenderer.DrawText(g, Text, Font, r, Enabled ? Fg : Color.FromArgb(148, 163, 184), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter); }
        protected override void OnMouseEnter(EventArgs e) { _h = true; Invalidate(); }
        protected override void OnMouseLeave(EventArgs e) { _h = false; Invalidate(); }
        static Color Lt(Color c, int a) => Color.FromArgb(Math.Min(255, Math.Max(0, c.R + a)), Math.Min(255, Math.Max(0, c.G + a)), Math.Min(255, Math.Max(0, c.B + a)));
        static GraphicsPath RR(Rectangle r, int d) { var p = new GraphicsPath(); int dd = d * 2; p.AddArc(r.X, r.Y, dd, dd, 180, 90); p.AddArc(r.Right - dd, r.Y, dd, dd, 270, 90); p.AddArc(r.Right - dd, r.Bottom - dd, dd, dd, 0, 90); p.AddArc(r.X, r.Bottom - dd, dd, dd, 90, 90); p.CloseFigure(); return p; }
    }

    internal sealed class SlimBar : Control
    {
        public int Value; public Color Bar = Color.FromArgb(99, 102, 241), Track = Color.FromArgb(226, 232, 240);
        public SlimBar() { Height = 6; SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true); }
        protected override void OnPaint(PaintEventArgs e) { var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias; int r = Height / 2; using (var b = new SolidBrush(Track)) using (var p = P(0, 0, Width, Height, r)) g.FillPath(b, p); int w = (int)(Width * Math.Max(0, Math.Min(100, Value)) / 100.0); if (w > 2) using (var b = new SolidBrush(Bar)) using (var p = P(0, 0, w, Height, r)) g.FillPath(b, p); }
        static GraphicsPath P(int x, int y, int w, int h, int r) { var p = new GraphicsPath(); if (w <= 0) return p; r = Math.Min(r, Math.Min(w / 2, h / 2)); int d = r * 2; p.AddArc(x, y, d, d, 180, 90); p.AddArc(x + w - d, y, d, d, 270, 90); p.AddArc(x + w - d, y + h - d, d, d, 0, 90); p.AddArc(x, y + h - d, d, d, 90, 90); p.CloseFigure(); return p; }
    }

    internal sealed class KpiCard : Panel
    {
        public string Title = "", Val = "", Sub = "";
        public Color ValColor = Color.FromArgb(15, 23, 42);
        public bool ShowButton;
        public event EventHandler ButtonClick;
        public KpiCard() { BackColor = Color.Transparent; Size = new Size(180, 80); SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var p = RR(new Rectangle(0, 0, Width - 1, Height - 1), 10))
            { using (var b = new SolidBrush(Color.White)) g.FillPath(b, p); using (var pen = new Pen(Color.FromArgb(226, 232, 240))) g.DrawPath(pen, p); }

            using (var tf = new Font("Segoe UI", 8.5f, FontStyle.Bold))
                TextRenderer.DrawText(g, Title, tf,
                    new Rectangle(4, 8, Width - 8, 16), Color.FromArgb(100, 116, 139),
                    TextFormatFlags.HorizontalCenter);

            if (ShowButton)
            {
                var btnRect = new Rectangle((Width - 64) / 2, 28, 64, 26);
                using (var bp = RR(btnRect, 6))
                using (var bb = new SolidBrush(Color.FromArgb(13, 148, 136)))
                { g.FillPath(bb, bp); }
                using (var bf = new Font("Segoe UI", 8.5f, FontStyle.Bold))
                    TextRenderer.DrawText(g, "🔍 조회", bf, btnRect, Color.White,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            else
            {
                float vfSize = 18f;
                if (Val != null)
                {
                    if (Val.Length > 14) vfSize = 9f;
                    else if (Val.Length > 10) vfSize = 10.5f;
                    else if (Val.Length > 6) vfSize = 12f;
                    else if (Val.Length > 4) vfSize = 14f;
                }
                using (var valFont = new Font("Segoe UI", vfSize, FontStyle.Bold))
                    TextRenderer.DrawText(g, Val, valFont,
                        new Rectangle(4, 24, Width - 8, 28), ValColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

                if (!string.IsNullOrEmpty(Sub))
                    using (var sf = new Font("Segoe UI", 7.5f))
                        TextRenderer.DrawText(g, Sub, sf,
                            new Rectangle(4, 56, Width - 8, 16), Color.FromArgb(148, 163, 184),
                            TextFormatFlags.HorizontalCenter | TextFormatFlags.EndEllipsis);
            }
        }
        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            if (ShowButton)
            {
                var btnRect = new Rectangle((Width - 64) / 2, 28, 64, 26);
                if (btnRect.Contains(e.Location))
                    ButtonClick?.Invoke(this, EventArgs.Empty);
            }
        }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (ShowButton)
            {
                var btnRect = new Rectangle((Width - 64) / 2, 28, 64, 26);
                Cursor = btnRect.Contains(e.Location) ? Cursors.Hand : Cursors.Default;
            }
            else Cursor = Cursors.Default;
        }
        static GraphicsPath RR(Rectangle r, int d) { var p = new GraphicsPath(); int dd = d * 2; p.AddArc(r.X, r.Y, dd, dd, 180, 90); p.AddArc(r.Right - dd, r.Y, dd, dd, 270, 90); p.AddArc(r.Right - dd, r.Bottom - dd, dd, dd, 0, 90); p.AddArc(r.X, r.Bottom - dd, dd, dd, 90, 90); p.CloseFigure(); return p; }
    }

    // ══════════════════════════════════════════════════════════
    //  MainForm
    // ══════════════════════════════════════════════════════════
    public partial class MainForm : Form
    {
        // ── 팔레트 ──
        static readonly Color MAIN_BG = Color.FromArgb(244, 246, 249);
        static readonly Color CARD_BG = Color.White;
        static readonly Color CARD_BRD = Color.FromArgb(226, 232, 240);
        static readonly Color HDR_TXT = Color.FromArgb(15, 23, 42);
        static readonly Color TEAL = Color.FromArgb(79, 70, 229);
        static readonly Color TEAL_D = Color.FromArgb(67, 56, 202);
        static readonly Color CORAL = Color.FromArgb(244, 63, 94);
        static readonly Color GREEN = Color.FromArgb(16, 185, 129);
        static readonly Color AMBER = Color.FromArgb(245, 158, 11);
        static readonly Color TXT_MAIN = Color.FromArgb(30, 41, 59);
        static readonly Color TXT_SEC = Color.FromArgb(71, 85, 105);
        static readonly Color TXT_MUTE = Color.FromArgb(148, 163, 184);
        static readonly Color GRID_HDR = Color.FromArgb(248, 250, 252);
        static readonly Color GRID_ALT = Color.White;
        static readonly Color GRID_LN = Color.FromArgb(238, 242, 246);
        static readonly Color GRID_SEL = Color.FromArgb(238, 242, 255);
        static readonly Color UP_RED = Color.FromArgb(153, 27, 27);      // 진한 붉은색(상승)
        static readonly Color DOWN_BLUE = Color.FromArgb(30, 64, 175);   // 진한 파란색(하락)

        // ── 상태 ──
        AxKHOpenAPI _ax;
        List<string> _codes = new List<string>();
        List<AnalysisResult> _res = new List<AnalysisResult>();
        List<SectorSupplySummary> _sK = new List<SectorSupplySummary>(), _sD = new List<SectorSupplySummary>();
        CancellationTokenSource _cts; bool _running;
        ConsensusClient _consensus = new ConsensusClient();
        AnalysisResult _selectedResult;
        System.Windows.Forms.Timer _liveRefreshTimer;
        bool _autoRefreshBusy;
        bool _autoReanalyzeEnabled = false;   // 기본 OFF (UI에서 ON/OFF)
        bool _intradayRefreshEnabled = false; // 기본 OFF (UI에서 ON/OFF)
        DateTime _lastAutoRefresh = DateTime.MinValue;
        static readonly int AutoRefreshIntervalMs = 180000; // 3분(장중 준실시간 재분석)
        static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "StockAnalyzer", "window.ini");

        // ── 컨트롤 ──
        RndBtn _btnLogin, _btnCsv, _btnRun, _btnStop, _btnSet;
        CheckBox _chkAutoReanalyze, _chkIntradayRefresh;
        ComboBox _cbCond; Label _lblLogin, _lblCsv, _lblProg;
        SlimBar _bar;
        DataGridView _gStock, _gResult, _gSector;
        Panel _pDetail;

        // ── 상단 종목 요약 KPI ──
        KpiCard _kpiScore, _kpiSupply, _kpiSectorSupply, _kpiConsensus;

        public MainForm() { InitializeComponent(); Load += (s, e) => { BuildOcx(); BuildUI(); }; }

        void BuildOcx()
        {
            try { _ax = new AxKHOpenAPI(); ((System.ComponentModel.ISupportInitialize)_ax).BeginInit(); _ax.Visible = false; _ax.Width = 1; _ax.Height = 1; Controls.Add(_ax); ((System.ComponentModel.ISupportInitialize)_ax).EndInit(); }
            catch (Exception ex) { MessageBox.Show("키움 OCX 오류:\n" + ex.Message); }
        }

        // ═══════════════ UI 빌드 ═══════════════════════════════

        void BuildUI()
        {
            Text = "Stock Analyzer"; MinimumSize = new Size(1100, 550);
            BackColor = MAIN_BG; ForeColor = TXT_MAIN;
            Font = new Font("Segoe UI", 9.5f);
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);

            RestoreWindowBounds();

            // 1. 하단 상태 표시줄
            var statusBar = new Panel { Dock = DockStyle.Bottom, Height = 32, BackColor = Color.White };
            statusBar.Paint += (s, e) => { using (var p = new Pen(GRID_LN)) e.Graphics.DrawLine(p, 0, 0, statusBar.Width, 0); };

            _bar = new SlimBar { Width = 180 };
            _lblProg = new Label { Text = "준비 완료", ForeColor = TXT_SEC, Font = new Font("Segoe UI", 9f, FontStyle.Bold), BackColor = Color.Transparent, TextAlign = ContentAlignment.MiddleLeft };
            statusBar.Controls.Add(_lblProg);
            statusBar.Controls.Add(_bar);
            statusBar.Resize += (s, e) => {
                _bar.Location = new Point(statusBar.Width - _bar.Width - 24, 13);
                _lblProg.SetBounds(24, 0, statusBar.Width - _bar.Width - 60, 32);
            };

            // 2. 통합 상단바
            var topBar = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = MAIN_BG };
            topBar.Paint += (s, e) => { using (var p = new Pen(CARD_BRD)) e.Graphics.DrawLine(p, 24, topBar.Height - 1, topBar.Width - 24, topBar.Height - 1); };

            var title = new Label { Text = "📊 Stock Analyzer", Font = new Font("Segoe UI", 15f, FontStyle.Bold), ForeColor = HDR_TXT, BackColor = MAIN_BG, TextAlign = ContentAlignment.MiddleLeft };

            _btnCsv = new RndBtn("📂 CSV", Color.White, TXT_MAIN, 90, 32) { Bdr = CARD_BRD, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold) };
            _btnCsv.Click += BtnCsv_Click;
            _lblCsv = new Label { Text = "선택된 파일 없음", ForeColor = TXT_MUTE, Font = new Font("Segoe UI", 8.5f), BackColor = MAIN_BG, TextAlign = ContentAlignment.MiddleLeft };

            var lcond = new Label { Text = "조건검색", ForeColor = TXT_SEC, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), BackColor = MAIN_BG, TextAlign = ContentAlignment.MiddleRight };
            _cbCond = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.White, ForeColor = TXT_MAIN, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f) };
            _cbCond.SelectedIndexChanged += CbCond_Changed;

            _btnRun = new RndBtn("▶ 분석 시작", TEAL, Color.White, 110, 32) { Enabled = false };
            _btnRun.Click += BtnRun_Click;
            _btnStop = new RndBtn("■ 중지", CORAL, Color.White, 76, 32) { Enabled = false };
            _btnStop.Click += (s, e) => _cts?.Cancel();

            _btnSet = new RndBtn("⚙ 설정", Color.White, TXT_MAIN, 70, 32) { Bdr = CARD_BRD };
            _btnSet.Click += (s, e) => { using (var f = new SettingsForm()) f.ShowDialog(this); };

            _chkAutoReanalyze = MkTopToggle("자동 재분석", _autoReanalyzeEnabled);
            _chkAutoReanalyze.CheckedChanged += (s, e) =>
            {
                _autoReanalyzeEnabled = _chkAutoReanalyze.Checked;
                ApplyRuntimeToggles();
                _lblProg.Text = _autoReanalyzeEnabled ? "자동 재분석: ON" : "자동 재분석: OFF";
            };

            _chkIntradayRefresh = MkTopToggle("장중 자동갱신", _intradayRefreshEnabled);
            _chkIntradayRefresh.CheckedChanged += (s, e) =>
            {
                _intradayRefreshEnabled = _chkIntradayRefresh.Checked;
                ApplyRuntimeToggles();
                _lblProg.Text = _intradayRefreshEnabled ? "장중 자동갱신: ON (다음 분석부터 반영)" : "장중 자동갱신: OFF";
            };

            _lblLogin = new Label { ForeColor = Color.FromArgb(244, 63, 94), Font = new Font("Segoe UI", 9f, FontStyle.Bold), BackColor = MAIN_BG, TextAlign = ContentAlignment.MiddleRight, Text = "● 미연결" };
            _btnLogin = new RndBtn("로그인", TEAL, Color.White, 76, 32);
            _btnLogin.Click += BtnLogin_Click;

            topBar.Controls.AddRange(new Control[] { title, _btnCsv, _lblCsv, lcond, _cbCond, _btnRun, _btnStop, _btnSet, _chkAutoReanalyze, _chkIntradayRefresh, _lblLogin, _btnLogin });

            topBar.Resize += (s, e) => {
                int y = 16;
                title.SetBounds(24, 0, 200, topBar.Height);

                int x = 224;
                _btnCsv.Location = new Point(x, y); x += _btnCsv.Width + 8;
                _lblCsv.SetBounds(x, y, 110, 32); x += 110;

                lcond.SetBounds(x, y, 60, 32); x += 64;
                _cbCond.SetBounds(x, y + 4, 160, 26); x += 172;

                _btnRun.Location = new Point(x, y); x += _btnRun.Width + 8;
                _btnStop.Location = new Point(x, y); x += _btnStop.Width + 8;
                _btnSet.Location = new Point(x, y); x += _btnSet.Width + 10;

                _chkAutoReanalyze.SetBounds(x, y + 4, 92, 24); x += 96;
                _chkIntradayRefresh.SetBounds(x, y + 4, 102, 24);

                _btnLogin.Location = new Point(topBar.Width - 100, y);
                _lblLogin.SetBounds(topBar.Width - 100 - 160, 0, 150, topBar.Height);
            };

            // 3. 본문 
            var body = new Panel { Dock = DockStyle.Fill, BackColor = MAIN_BG, Padding = new Padding(24, 12, 24, 24) };

            var kpiRow = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 96, BackColor = MAIN_BG, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(0, 0, 0, 12) };
            _kpiScore = new KpiCard { Title = "총점", Val = "—", Sub = "종목을 선택하세요" };
            _kpiSupply = new KpiCard { Title = "수급", Val = "—", Sub = "" };
            _kpiSectorSupply = new KpiCard { Title = "업종수급", Val = "—", Sub = "" };
            _kpiConsensus = new KpiCard { Title = "컨센서스", Val = "—", Sub = "", ShowButton = false };
            _kpiConsensus.ButtonClick += BtnConsensus_Click;
            kpiRow.Controls.AddRange(new Control[] { _kpiScore, Sp(16), _kpiSupply, Sp(16), _kpiSectorSupply, Sp(16), _kpiConsensus });

            var grid3 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = MAIN_BG, Padding = new Padding(0) };
            grid3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
            grid3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            grid3.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            BuildP1();
            grid3.Controls.Add(WCard("ANALYSIS RESULT", BuildP2()), 0, 0);
            grid3.Controls.Add(BuildP3(), 1, 0);

            body.Controls.Add(grid3);
            body.Controls.Add(kpiRow);

            Controls.Add(body);
            Controls.Add(topBar);
            Controls.Add(statusBar);
            InitAutoRefreshTimer();
            ApplyRuntimeToggles();
        }

        // ── 패널 빌더 ──

        Control BuildP1()
        {
            _gStock = MkGrid(("종목명", "Name", 110), ("코드", "Code", 60), ("시장", "Market", 45));
            _gStock.SelectionChanged += GSel;
            return _gStock;
        }

        Control BuildP2()
        {
            // ★ 새 그리드: 3단 헤더 (수급(수량) → 외국인/기관 → 5일/10일)
            _gResult = MkGrid(
                ("순위", "Rank", 32),
                ("종목", "Name", 78),
                ("총점", "TotalScore", 40),
                ("5일", "F5D", 56),         // 외국인 5일
                ("10일", "F10D", 56),        // 외국인 10일
                ("5일", "I5D", 56),          // 기관 5일
                ("10일", "I10D", 56),        // 기관 10일
                ("추세", "Trend", 42),
                ("거래량\n(20/60)", "Vol", 52),
                ("업종", "Sector", 60));

            _gResult.ColumnHeadersHeight = 54;
            _gResult.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            _gResult.RowTemplate.Height = 30;

            _gResult.SelectionChanged += GRSel;
            _gResult.CellFormatting += GRFmt;
            _gResult.CellPainting += GRCellPaint;   // 하단 5일/10일 텍스트 위치
            _gResult.Paint += GRHeaderPaint;          // 상단 그룹 라벨 오버레이
            return _gResult;
        }

        Control BuildP3()
        {
            var outer = new Panel { Dock = DockStyle.Fill, BackColor = MAIN_BG, Padding = new Padding(8, 0, 0, 0) };
            var sp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = MAIN_BG, CellBorderStyle = TableLayoutPanelCellBorderStyle.None };
            sp.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            sp.RowStyles.Add(new RowStyle(SizeType.Percent, 60));

            _gSector = MkGrid(("업종", "SectorName", 80), ("시장", "Market", 40),
                ("외국인", "ForeignNet5DB", 65), ("기관", "InstNet5DB", 65), ("합산", "TotalNet5DB", 65));
            _gSector.CellFormatting += SectorFmt;
            sp.Controls.Add(WCard("SECTOR TREND", _gSector, 0, 0, 0, 4), 0, 0);

            _pDetail = new Panel { Dock = DockStyle.Fill, BackColor = CARD_BG, AutoScroll = true };
            ShowDetail(null);
            sp.Controls.Add(WCard("STOCK DETAIL", _pDetail, 0, 0, 0, 0), 0, 1);

            outer.Controls.Add(sp);
            return outer;
        }

        static Panel WCard(string title, Control inner, int l = 0, int t = 0, int r = 12, int b = 0)
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = MAIN_BG, Padding = new Padding(l, t, r, b) };
            var card = new Panel { Dock = DockStyle.Fill, BackColor = CARD_BG, Padding = new Padding(1) };
            var hdr = new Panel { Dock = DockStyle.Top, Height = 36, BackColor = CARD_BG };
            hdr.Paint += (s, e) => { TextRenderer.DrawText(e.Graphics, title, new Font("Segoe UI", 9f, FontStyle.Bold), new Point(12, 10), TXT_SEC); using (var pen = new Pen(Color.FromArgb(238, 242, 246))) e.Graphics.DrawLine(pen, 0, 35, hdr.Width, 35); };
            card.Controls.Add(inner);
            card.Controls.Add(hdr);
            p.Controls.Add(card);
            return p;
        }

        CheckBox MkTopToggle(string text, bool initial)
        {
            var cb = new CheckBox
            {
                Appearance = Appearance.Button,
                AutoSize = false,
                Text = text,
                TextAlign = ContentAlignment.MiddleCenter,
                Checked = initial,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = TXT_SEC,
                Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0)
            };
            cb.FlatAppearance.BorderSize = 1;
            cb.FlatAppearance.CheckedBackColor = Color.White;
            cb.FlatAppearance.MouseDownBackColor = Color.White;
            cb.FlatAppearance.MouseOverBackColor = Color.White;
            cb.CheckedChanged += AutoToggleStyleChanged;
            ApplyToggleStyle(cb);
            return cb;
        }

        void ApplyRuntimeToggles()
        {
            if (_chkAutoReanalyze != null && _chkAutoReanalyze.Checked != _autoReanalyzeEnabled)
                _chkAutoReanalyze.Checked = _autoReanalyzeEnabled;
            if (_chkIntradayRefresh != null && _chkIntradayRefresh.Checked != _intradayRefreshEnabled)
                _chkIntradayRefresh.Checked = _intradayRefreshEnabled;

            if (_chkAutoReanalyze != null) ApplyToggleStyle(_chkAutoReanalyze);
            if (_chkIntradayRefresh != null) ApplyToggleStyle(_chkIntradayRefresh);

            if (_liveRefreshTimer != null)
                _liveRefreshTimer.Enabled = _autoReanalyzeEnabled && !_running && _codes != null && _codes.Count > 0;
        }

        void AutoToggleStyleChanged(object sender, EventArgs e)
        {
            if (sender is CheckBox cb) ApplyToggleStyle(cb);
        }

        void ApplyToggleStyle(CheckBox cb)
        {
            if (cb == null) return;
            cb.FlatAppearance.BorderColor = cb.Checked ? TEAL_D : CARD_BRD;
            cb.BackColor = cb.Checked ? Color.FromArgb(238, 242, 255) : Color.White;
            cb.ForeColor = cb.Checked ? TEAL_D : TXT_SEC;
        }

        // ═══════════════ 이벤트 핸들러 ═══════════════════════════

        async void BtnLogin_Click(object s, EventArgs e)
        {
            try
            {
                using (var k = new KiwoomClient(_ax))
                {
                    _lblLogin.Text = "⏳ 로그인 중..."; _lblLogin.ForeColor = AMBER;
                    await k.LoginAsync();
                    _lblLogin.Text = "● 연결됨"; _lblLogin.ForeColor = GREEN;
                    await LoadConditions();
                }
            }
            catch (Exception ex) { _lblLogin.Text = $"✗ {ex.Message}"; _lblLogin.ForeColor = CORAL; }
        }

        async Task LoadConditions()
        {
            try
            {
                using (var k = new KiwoomClient(_ax))
                {
                    var conds = await k.GetConditionListAsync();
                    _cbCond.Items.Clear(); _cbCond.Items.Add("-- 조건식 선택 --");
                    foreach (var c in conds) _cbCond.Items.Add(new CI(c.idx, c.name));
                    _cbCond.SelectedIndex = 0;
                }
            }
            catch { }
        }

        async void CbCond_Changed(object s, EventArgs e)
        {
            if (_cbCond.SelectedItem is CI ci)
            {
                try
                {
                    using (var k = new KiwoomClient(_ax))
                    {
                        var condCodes = await k.GetConditionCodesAsync(ci.Idx, ci.Nm);
                        _codes = condCodes;
                        _lblCsv.Text = $"조건식: {ci.Nm} ({_codes.Count}개)";
                        _btnRun.Enabled = _codes.Count > 0 && !_running;
                    }
                }
                catch (Exception ex) { MessageBox.Show("조건검색 오류: " + ex.Message); }
            }
        }

        void BtnCsv_Click(object s, EventArgs e)
        {
            using (var d = new OpenFileDialog { Filter = "CSV|*.csv", Title = "종목 코드 CSV" })
            {
                if (d.ShowDialog() != DialogResult.OK) return;
                _codes = CsvCodeExtractor.Extract(d.FileName);
                _lblCsv.Text = $"{Path.GetFileName(d.FileName)} ({_codes.Count}개)";
                _btnRun.Enabled = _codes.Count > 0 && !_running;
            }
        }

        void GSel(object s, EventArgs e) { }

        void GRSel(object s, EventArgs e)
        {
            if (_gResult.SelectedRows.Count == 0) return;
            var idx = _gResult.SelectedRows[0].Index;
            if (idx < 0 || idx >= _res.Count) return;
            var r = _res[idx];
            ShowDetail(r);
            UpdateStockSummary(r);
        }

        async void BtnRun_Click(object s, EventArgs e)
        {
            await RunAnalysisCoreAsync(false);
        }

        void InitAutoRefreshTimer()
        {
            if (_liveRefreshTimer != null) return;
            _liveRefreshTimer = new System.Windows.Forms.Timer { Interval = AutoRefreshIntervalMs };
            _liveRefreshTimer.Tick += async (s, e) =>
            {
                if (!_autoReanalyzeEnabled) return;
                if (_autoRefreshBusy || _running) return;
                if (_codes == null || _codes.Count == 0) return;
                if (!IsMarketHoursApprox()) return;
                if ((DateTime.Now - _lastAutoRefresh).TotalSeconds < 60) return;

                _autoRefreshBusy = true;
                try { await RunAnalysisCoreAsync(true); }
                finally { _autoRefreshBusy = false; }
            };
        }

        static bool IsMarketHoursApprox()
        {
            var now = DateTime.Now;
            var t = now.TimeOfDay;
            if (t < new TimeSpan(8, 55, 0) || t > new TimeSpan(15, 45, 0)) return false;
            return TradingDayHelper.IsTradingDay(now.Date);
        }

        async Task RunAnalysisCoreAsync(bool isAutoRefresh)
        {
            if (_running) return;
            if (_codes == null || _codes.Count == 0) return;

            _cts = new CancellationTokenSource();
            SetRun(true);

            if (!isAutoRefresh)
            {
                _res.Clear(); _sK.Clear(); _sD.Clear();
                _gResult.Rows.Clear(); _gSector.Rows.Clear();
                ShowDetail(null);
            }

            try
            {
                var cfg = ScoreConfig.Load();
                var engine = new AnalysisEngine(_ax);
                engine.Log += msg => { _lblProg.Text = (isAutoRefresh ? "[자동갱신] " : "") + msg; System.Diagnostics.Debug.WriteLine(msg); };
                engine.Progress += (cur, tot, nm) =>
                {
                    _bar.Value = tot <= 0 ? 0 : cur * 100 / tot; _bar.Invalidate();
                    _lblProg.Text = $"{(isAutoRefresh ? "자동갱신" : nm)} ({cur}/{tot})";
                };

                var (results, sK, sD) = await engine.RunAsync(_codes, cfg, _cts.Token, intradayRefresh: _intradayRefreshEnabled);

                _res = results; _sK = sK; _sD = sD;
                FillResult(); FillSector(); FillStock();
                UpdateKpi();

                _lastAutoRefresh = DateTime.Now;
                if (_liveRefreshTimer != null) _liveRefreshTimer.Enabled = _autoReanalyzeEnabled;
                if (isAutoRefresh) _lblProg.Text = $"자동갱신 완료 {DateTime.Now:HH:mm:ss}";
            }
            catch (OperationCanceledException) { _lblProg.Text = "취소됨"; }
            catch (Exception ex)
            {
                _lblProg.Text = $"오류: {ex.Message}";
                if (!isAutoRefresh) MessageBox.Show(ex.Message, "오류");
            }
            finally { SetRun(false); }
        }

        // ── 그리드 포맷팅 ──

        void GRFmt(object s, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _res.Count) return;
            var r = _res[e.RowIndex]; var col = _gResult.Columns[e.ColumnIndex].Name;

            if (col == "TotalScore")
                e.CellStyle.ForeColor = r.TotalScore >= 80 ? GREEN : r.TotalScore >= 50 ? TEAL_D : r.TotalScore >= 30 ? AMBER : CORAL;

            // 외국인/기관 수량(환산) 색상: 상승(+) 진한 빨강 / 하락(-) 진한 파랑 + 굵게
            if (col == "F5D" || col == "F10D" || col == "I5D" || col == "I10D")
            {
                var v = e.Value?.ToString() ?? "";
                if (v.StartsWith("+")) e.CellStyle.ForeColor = UP_RED;
                else if (v.StartsWith("-")) e.CellStyle.ForeColor = DOWN_BLUE;
                else e.CellStyle.ForeColor = TXT_MUTE;
                e.CellStyle.Font = new Font("Segoe UI", 7.8f, FontStyle.Bold);
            }

            if (col == "Trend")
            {
                if (r.SupplyTrend == SupplyTrend.상승 || r.SupplyTrend == SupplyTrend.상승반전) e.CellStyle.ForeColor = UP_RED;
                else if (r.SupplyTrend == SupplyTrend.하락 || r.SupplyTrend == SupplyTrend.하락반전) e.CellStyle.ForeColor = DOWN_BLUE;
                else e.CellStyle.ForeColor = TXT_MUTE;
                e.CellStyle.Font = new Font("Segoe UI", 8.2f, FontStyle.Bold);
            }

            if (col == "Vol")
            {
                if (r.VolTrend == VolTrend.상승) e.CellStyle.ForeColor = UP_RED;
                else if (r.VolTrend == VolTrend.하락) e.CellStyle.ForeColor = DOWN_BLUE;
                else e.CellStyle.ForeColor = TXT_MUTE;
                e.CellStyle.Font = new Font("Segoe UI", 8.2f, FontStyle.Bold);
            }
        }

        void SectorFmt(object s, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var row = _gSector.Rows[e.RowIndex];
            var sectorName = row.Cells["SectorName"]?.Value?.ToString() ?? "";
            bool isGroupRow = sectorName.StartsWith("■ ");

            if (isGroupRow)
            {
                e.CellStyle.BackColor = Color.FromArgb(248, 250, 252);
                e.CellStyle.SelectionBackColor = e.CellStyle.BackColor;
                e.CellStyle.ForeColor = TXT_SEC;
                e.CellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);
                return;
            }

            var col = _gSector.Columns[e.ColumnIndex].Name;
            if (col == "ForeignNet5DB" || col == "InstNet5DB" || col == "TotalNet5DB")
            {
                var v = e.Value?.ToString() ?? "";
                if (v.StartsWith("+")) e.CellStyle.ForeColor = UP_RED;
                else if (v.StartsWith("-")) e.CellStyle.ForeColor = DOWN_BLUE;
                else e.CellStyle.ForeColor = TXT_MUTE;
                e.CellStyle.Font = new Font("Segoe UI", 8.2f, FontStyle.Bold);
            }
        }

        /// <summary>헤더 하단 "5일"/"10일" 텍스트를 하단 1/3에 배치, "추세"는 오버레이에서 처리</summary>
        void GRCellPaint(object s, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex != -1) return; // 헤더만
            var col = _gResult.Columns[e.ColumnIndex].Name;
            if (col != "F5D" && col != "F10D" && col != "I5D" && col != "I10D" && col != "Trend") return;

            e.Handled = true;
            var rect = e.CellBounds;

            // 배경 + 테두리
            using (var bg = new SolidBrush(GRID_HDR))
                e.Graphics.FillRectangle(bg, rect);
            using (var pen = new Pen(GRID_LN))
            {
                e.Graphics.DrawLine(pen, rect.Right - 1, rect.Top, rect.Right - 1, rect.Bottom);
                e.Graphics.DrawLine(pen, rect.Left, rect.Bottom - 1, rect.Right, rect.Bottom - 1);
            }

            if (col == "Trend") return; // "추세"는 Paint 오버레이에서 통합 그림

            // "5일"/"10일"을 하단 1/3에 배치
            int row3Y = rect.Y + (rect.Height * 2 / 3);
            int row3H = rect.Height - (rect.Height * 2 / 3);
            var r3 = new Rectangle(rect.X, row3Y, rect.Width, row3H);

            string txt = (col == "F5D" || col == "I5D") ? "5일" : "10일";
            using (var f = new Font("Segoe UI", 7.5f, FontStyle.Bold))
                TextRenderer.DrawText(e.Graphics, txt, f, r3, TXT_SEC,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        /// <summary>3단 그룹 헤더를 Paint 오버레이로 그림 (CellPainting은 셀 바깥 클리핑됨)</summary>
        void GRHeaderPaint(object s, PaintEventArgs e)
        {
            if (_gResult.Columns.Count < 7) return;

            var g = e.Graphics;
            // 컬럼 좌표 계산 (스크롤 고려)
            int GetColLeft(string name)
            {
                int x = _gResult.RowHeadersVisible ? _gResult.RowHeadersWidth : 0;
                for (int i = 0; i < _gResult.Columns.Count; i++)
                {
                    if (_gResult.Columns[i].Name == name) return x - _gResult.HorizontalScrollingOffset;
                    if (_gResult.Columns[i].Visible) x += _gResult.Columns[i].Width;
                }
                return x;
            }
            int GetColW(string name) => _gResult.Columns[name].Visible ? _gResult.Columns[name].Width : 0;

            int hdrH = _gResult.ColumnHeadersHeight;
            int row1H = hdrH / 3;
            int row2H = hdrH / 3;
            int row3H = hdrH - row1H - row2H;

            int xF5D = GetColLeft("F5D");
            int xI5D = GetColLeft("I5D");
            int xTrend = GetColLeft("Trend");
            int wF = GetColW("F5D") + GetColW("F10D");
            int wI = GetColW("I5D") + GetColW("I10D");
            int wTrend = GetColW("Trend");
            int totalW = wF + wI + wTrend;

            var hdrFont = new Font("Segoe UI", 7.5f, FontStyle.Bold);
            var grpFont = new Font("Segoe UI", 8f, FontStyle.Bold);

            // Row 1: "수급(수량)" 전체 그룹
            var grpRect = new Rectangle(xF5D, 0, totalW, row1H);
            using (var bg = new SolidBrush(GRID_HDR)) g.FillRectangle(bg, grpRect);
            TextRenderer.DrawText(g, "수급(수량·1주)", grpFont, grpRect, TXT_SEC,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            using (var pen = new Pen(GRID_LN))
                g.DrawLine(pen, xF5D, row1H, xF5D + totalW, row1H);

            // Row 2: "외국인" | "기관" | "추세"
            var fRect = new Rectangle(xF5D, row1H, wF, row2H);
            using (var bg = new SolidBrush(GRID_HDR)) g.FillRectangle(bg, fRect);
            TextRenderer.DrawText(g, "외국인", hdrFont, fRect, TXT_SEC,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

            var iRect = new Rectangle(xI5D, row1H, wI, row2H);
            using (var bg = new SolidBrush(GRID_HDR)) g.FillRectangle(bg, iRect);
            TextRenderer.DrawText(g, "기관", hdrFont, iRect, TXT_SEC,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

            // "추세"는 row2+row3 통합
            var tRect = new Rectangle(xTrend, row1H, wTrend, row2H + row3H);
            using (var bg = new SolidBrush(GRID_HDR)) g.FillRectangle(bg, tRect);
            TextRenderer.DrawText(g, "추세", hdrFont, tRect, TXT_SEC,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

            // 구분선
            using (var pen = new Pen(GRID_LN))
            {
                g.DrawLine(pen, xF5D, row1H + row2H, xF5D + wF, row1H + row2H);
                g.DrawLine(pen, xI5D, row1H + row2H, xI5D + wI, row1H + row2H);
                // 세로 구분선
                g.DrawLine(pen, xI5D, row1H, xI5D, row1H + row2H);
                g.DrawLine(pen, xTrend, row1H, xTrend, hdrH);
            }

            hdrFont.Dispose();
            grpFont.Dispose();
        }

        // ── 데이터 갱신 ──

        void FillResult()
        {
            _gResult.Rows.Clear();
            for (int i = 0; i < _res.Count; i++)
            {
                var r = _res[i];
                _gResult.Rows.Add(
                    i + 1,
                    r.Name,
                    r.TotalScore.ToString("F1"),
                    FQ1(r.ForeignNet5D),    // F5D (실제 수량, 주)
                    FQ1(r.ForeignNet10D),   // F10D
                    FQ1(r.InstNet5D),       // I5D
                    FQ1(r.InstNet10D),      // I10D
                    r.SupplyTrend.ToString(),             // Trend
                    r.VolTrend.ToString(),                // Vol (20/60)
                    r.SectorName,                         // Sector
                    r.Code);
            }
        }

        void FillSector()
        {
            _gSector.Rows.Clear();

            void AddSectorBlock(string title, IEnumerable<SectorSupplySummary> items)
            {
                var list = (items ?? Enumerable.Empty<SectorSupplySummary>()).ToList();
                if (list.Count == 0) return;

                int hdrIdx = _gSector.Rows.Add("■ " + title, "", "", "", "");
                var hdrRow = _gSector.Rows[hdrIdx];
                hdrRow.DefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);
                hdrRow.DefaultCellStyle.SelectionBackColor = hdrRow.DefaultCellStyle.BackColor;
                hdrRow.DefaultCellStyle.ForeColor = TXT_SEC;
                hdrRow.DefaultCellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);

                foreach (var x in list.OrderByDescending(x => x.TotalNet5D))
                    _gSector.Rows.Add(x.SectorName, x.Market, FA(x.ForeignNet5D), FA(x.InstNet5D), FA(x.TotalNet5D));

                _gSector.Rows.Add("", "", "", "", ""); // 시각 구분용 빈 줄
                var gap = _gSector.Rows[_gSector.Rows.Count - 1];
                gap.DefaultCellStyle.BackColor = CARD_BG;
                gap.DefaultCellStyle.SelectionBackColor = CARD_BG;
                gap.Height = 6;
            }

            AddSectorBlock("KOSPI", _sK);
            AddSectorBlock("KOSDAQ", _sD);

            // 마지막 빈 줄 제거 (있으면)
            if (_gSector.Rows.Count > 0)
            {
                var last = _gSector.Rows[_gSector.Rows.Count - 1];
                bool isBlank = string.IsNullOrWhiteSpace(Convert.ToString(last.Cells["SectorName"].Value))
                            && string.IsNullOrWhiteSpace(Convert.ToString(last.Cells["Market"].Value));
                if (isBlank) _gSector.Rows.RemoveAt(_gSector.Rows.Count - 1);
            }
        }

        void FillStock()
        {
            _gStock.Rows.Clear();
            foreach (var c in _codes)
            {
                var r = _res.FirstOrDefault(x => x.Code == c);
                _gStock.Rows.Add(r?.Name ?? c, c, r?.Market ?? "");
            }
        }

        void UpdateKpi()
        {
            if (_res.Count > 0 && _gResult.Rows.Count > 0)
            { _gResult.ClearSelection(); _gResult.Rows[0].Selected = true; }
        }

        void UpdateStockSummary(AnalysisResult r)
        {
            _selectedResult = r;

            if (r == null)
            {
                _kpiScore.Val = "—"; _kpiScore.Sub = "종목을 선택하세요"; _kpiScore.ValColor = TXT_MAIN;
                _kpiSupply.Val = "—"; _kpiSupply.Sub = ""; _kpiSupply.ValColor = TXT_MAIN;
                _kpiSectorSupply.Val = "—"; _kpiSectorSupply.Sub = ""; _kpiSectorSupply.ValColor = TXT_MAIN;
                _kpiConsensus.Val = "—"; _kpiConsensus.Sub = ""; _kpiConsensus.ValColor = TXT_MAIN;
                _kpiConsensus.ShowButton = false;
            }
            else
            {
                _kpiScore.Val = r.TotalScore.ToString("F1"); _kpiScore.Sub = $"{r.Name} ({r.Market})"; _kpiScore.ValColor = SC(r.TotalScore);

                var supGrade = SupplyGrade(r.StockSupplyScore);
                _kpiSupply.Val = supGrade.text; _kpiSupply.Sub = $"종목수급 {r.StockSupplyScore:F1}점"; _kpiSupply.ValColor = supGrade.color;

                var secGrade = SupplyGrade(r.SectorSupplyScore);
                _kpiSectorSupply.Val = secGrade.text; _kpiSectorSupply.Sub = $"업종수급 {r.SectorSupplyScore:F1}점"; _kpiSectorSupply.ValColor = secGrade.color;

                _kpiConsensus.Val = ""; _kpiConsensus.Sub = ""; _kpiConsensus.ShowButton = true;
            }
            _kpiScore.Invalidate(); _kpiSupply.Invalidate(); _kpiSectorSupply.Invalidate(); _kpiConsensus.Invalidate();
        }

        /// <summary>컨센서스 조회 버튼 클릭</summary>
        async void BtnConsensus_Click(object s, EventArgs e)
        {
            if (_selectedResult == null) return;

            var r = _selectedResult;
            _kpiConsensus.ShowButton = false;
            _kpiConsensus.Val = "⏳"; _kpiConsensus.Sub = "조회중..."; _kpiConsensus.ValColor = AMBER;
            _kpiConsensus.Invalidate();

            try
            {
                var data = await _consensus.GetConsensusAsync(r.Code);

                if (data != null && data.IsValid)
                {
                    _kpiConsensus.Val = data.Opinion ?? "—";
                    var parts = new List<string>();
                    if (data.TargetPrice.HasValue) parts.Add($"목표가 {data.TargetPrice.Value:N0}원");
                    if (data.DeviationPct.HasValue) parts.Add($"괴리 {data.DeviationPct.Value:+0.0;-0.0}%");
                    _kpiConsensus.Sub = string.Join(" · ", parts);
                    _kpiConsensus.ValColor = ConsensusColor(data.Opinion);

                    ShowDetail(r, data);
                }
                else
                {
                    _kpiConsensus.Val = "데이터 없음"; _kpiConsensus.Sub = ""; _kpiConsensus.ValColor = TXT_MUTE;
                }
            }
            catch (Exception ex)
            {
                _kpiConsensus.Val = "오류"; _kpiConsensus.Sub = ex.Message; _kpiConsensus.ValColor = CORAL;
            }
            _kpiConsensus.ShowButton = false;
            _kpiConsensus.Invalidate();
        }

        static Color ConsensusColor(string opinion)
        {
            if (string.IsNullOrEmpty(opinion)) return TXT_MUTE;
            if (opinion.Contains("강력매수")) return Color.FromArgb(16, 185, 129);
            if (opinion.Contains("매수")) return GREEN;
            if (opinion.Contains("보유") || opinion.Contains("중립")) return AMBER;
            if (opinion.Contains("강력매도")) return Color.FromArgb(220, 38, 38);
            if (opinion.Contains("매도")) return CORAL;
            return TXT_MAIN;
        }

        static (string text, Color color) SupplyGrade(double score)
        {
            if (score >= 80) return ("아주좋음", GREEN);
            if (score >= 60) return ("좋음", Color.FromArgb(34, 197, 94));
            if (score >= 40) return ("보통", AMBER);
            if (score >= 20) return ("나쁨", Color.FromArgb(249, 115, 22));
            return ("아주나쁨", CORAL);
        }

        // ── 세부 패널 ──

        void ShowDetail(AnalysisResult r, ConsensusData con = null)
        {
            _pDetail.Controls.Clear();
            if (r == null) { _pDetail.Controls.Add(new Label { Text = "종목을 선택하세요", ForeColor = TXT_MUTE, Font = new Font("Segoe UI", 9.5f), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.White }); return; }

            var fl = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, BackColor = Color.White, Padding = new Padding(16, 12, 16, 12) };

            fl.Controls.Add(DL($"{r.Name}", new Font("Segoe UI", 14f, FontStyle.Bold), TXT_MAIN));
            fl.Controls.Add(DL($"{r.Code}  ·  {r.Market}  ·  {r.SectorName}  ·  {r.CurrentPrice:N0}원", new Font("Segoe UI", 8.5f), TXT_SEC));
            fl.Controls.Add(DH());

            fl.Controls.Add(DR("총점", r.TotalScore.ToString("F1"), SC(r.TotalScore), true));
            fl.Controls.Add(DR("기업가치", r.ValueScore.ToString("F1"), TEAL_D));
            fl.Controls.Add(DR("종목수급", r.StockSupplyScore.ToString("F1"), TEAL_D));
            fl.Controls.Add(DR("업종수급", r.SectorSupplyScore.ToString("F1"), TEAL_D));
            fl.Controls.Add(DH());

            fl.Controls.Add(DR("PER", r.Per.HasValue ? r.Per.Value.ToString("F2") : "—"));
            fl.Controls.Add(DR("PBR", r.Pbr.HasValue ? r.Pbr.Value.ToString("F2") : "—"));
            fl.Controls.Add(DR("ROE", r.Roe.HasValue ? r.Roe.Value.ToString("F1") + "%" : "—"));
            fl.Controls.Add(DR("업종PER", r.SectorAvgPer.HasValue ? r.SectorAvgPer.Value.ToString("F2") : "—"));
            fl.Controls.Add(DR("업종PBR", r.SectorAvgPbr.HasValue ? r.SectorAvgPbr.Value.ToString("F2") : "—"));
            fl.Controls.Add(DH());

            // ★ 종목수급: 실제 수량(주)만 표시 (금액 제거)
            fl.Controls.Add(DR("외국인 당일", FQ1(r.ForeignNetD1), NQ(r.ForeignNetD1), true));
            fl.Controls.Add(DR("외국인 5일", FQ1(r.ForeignNet5D), NQ(r.ForeignNet5D), true));
            fl.Controls.Add(DR("외국인 10일", FQ1(r.ForeignNet10D), NQ(r.ForeignNet10D), true));
            fl.Controls.Add(DR("외국인 20일", FQ1(r.ForeignNet20D), NQ(r.ForeignNet20D), true));
            fl.Controls.Add(DR("기관 당일", FQ1(r.InstNetD1), NQ(r.InstNetD1), true));
            fl.Controls.Add(DR("기관 5일", FQ1(r.InstNet5D), NQ(r.InstNet5D), true));
            fl.Controls.Add(DR("기관 10일", FQ1(r.InstNet10D), NQ(r.InstNet10D), true));
            fl.Controls.Add(DR("기관 20일", FQ1(r.InstNet20D), NQ(r.InstNet20D), true));
            fl.Controls.Add(DH());

            fl.Controls.Add(DR("거래량20D평균", FVol(r.VolAvg20D)));
            fl.Controls.Add(DR("거래량60D평균", FVol(r.VolAvg60D)));
            fl.Controls.Add(DR("거래량추세", r.VolTrend.ToString(), r.VolTrend == VolTrend.상승 ? UP_RED : r.VolTrend == VolTrend.하락 ? DOWN_BLUE : TXT_SEC, true));
            fl.Controls.Add(DR("수급추세", r.SupplyTrend.ToString(), r.SupplyTrend == SupplyTrend.상승 || r.SupplyTrend == SupplyTrend.상승반전 ? UP_RED : (r.SupplyTrend == SupplyTrend.하락 || r.SupplyTrend == SupplyTrend.하락반전) ? DOWN_BLUE : AMBER, true));

            // 컨센서스 섹션
            if (con != null && con.IsValid)
            {
                fl.Controls.Add(DH());
                fl.Controls.Add(DR("컨센서스", con.Opinion ?? "—", ConsensusColor(con.Opinion), true));
                if (con.TargetPrice.HasValue) fl.Controls.Add(DR("목표가", $"{con.TargetPrice.Value:N0}원"));
                if (con.TargetPriceMin.HasValue && con.TargetPriceMax.HasValue) fl.Controls.Add(DR("목표가범위", $"{con.TargetPriceMin.Value:N0} ~ {con.TargetPriceMax.Value:N0}원"));
                if (con.DeviationPct.HasValue) fl.Controls.Add(DR("괴리율", $"{con.DeviationPct.Value:+0.0;-0.0}%", con.DeviationPct > 0 ? UP_RED : DOWN_BLUE, true));
                if (con.ConsensusPer.HasValue) fl.Controls.Add(DR("컨센PER", con.ConsensusPer.Value.ToString("F2")));
                if (con.ConsensusEps.HasValue) fl.Controls.Add(DR("컨센EPS", $"{con.ConsensusEps.Value:N0}원"));
                fl.Controls.Add(DR("애널리스트", $"{con.AnalystCount}명"));
                if (!string.IsNullOrEmpty(con.LatestReportDate)) fl.Controls.Add(DR("최신리포트", con.LatestReportDate));
                fl.Controls.Add(DR("출처", con.Source ?? "—"));
            }

            _pDetail.Controls.Add(fl);
        }

        // ═══════════════ 팩토리 ═══════════════════════════════

        static DataGridView MkGrid(params (string h, string n, int w)[] cols)
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = CARD_BG,
                BorderStyle = BorderStyle.None,
                GridColor = GRID_LN,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeColumns = false,
                AllowUserToResizeRows = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeight = 36,
                RowTemplate = { Height = 32 },
                EnableHeadersVisualStyles = false,
                ScrollBars = ScrollBars.Vertical,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = CARD_BG, ForeColor = TXT_MAIN, SelectionBackColor = GRID_SEL, SelectionForeColor = TXT_MAIN, Font = new Font("Segoe UI", 8.5f), Padding = new Padding(4, 0, 4, 0), Alignment = DataGridViewContentAlignment.MiddleCenter },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = GRID_HDR, ForeColor = TXT_SEC, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), SelectionBackColor = GRID_HDR, Alignment = DataGridViewContentAlignment.MiddleCenter, Padding = new Padding(0) },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = GRID_ALT, ForeColor = TXT_MAIN },
            };
            foreach (var (h, n, w) in cols)
                g.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = h, Name = n, MinimumWidth = w });
            if (cols.Any(c => c.n == "Name" && cols.Any(x => x.n == "TotalScore")))
                g.Columns.Add(new DataGridViewTextBoxColumn { Name = "Code2", Visible = false });
            return g;
        }

        static Panel Sp(int w) => new Panel { Width = w, Height = 1, BackColor = MAIN_BG };
        static Label DL(string t, Font f, Color c) => new Label { Text = t, AutoSize = true, Font = f, ForeColor = c, BackColor = CARD_BG, Margin = new Padding(0, 0, 0, 4) };
        static Panel DH() => new Panel { Width = 340, Height = 1, BackColor = GRID_LN, Margin = new Padding(0, 8, 0, 8) };
        static Panel DR(string lbl, string val, Color? vc = null, bool boldVal = false)
        {
            var p = new Panel { Width = 340, Height = 24, BackColor = CARD_BG };
            p.Controls.Add(new Label { Text = lbl, Width = 110, ForeColor = TXT_SEC, Font = new Font("Segoe UI", 8.5f), TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Left, BackColor = CARD_BG });
            p.Controls.Add(new Label { Text = val, ForeColor = vc ?? TXT_MAIN, Font = new Font("Segoe UI", 9f, boldVal ? FontStyle.Bold : FontStyle.Regular), TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill, BackColor = CARD_BG });
            return p;
        }

        // ── 포맷터 ──

        /// <summary>수량(주) 포맷: +12.3만, -5,678, +1.2억</summary>
        static string FQ(long v)
        {
            double a = Math.Abs(v);
            string sign = v >= 0 ? "+" : "-";
            if (a >= 1e8) return sign + (a / 1e8).ToString("F1") + "억";
            if (a >= 1e4) return sign + (a / 1e4).ToString("F1") + "만";
            if (a == 0) return "0";
            return v.ToString("+#,0;-#,0");
        }

        /// <summary>수량(주) 포맷: 1주 단위 표시(축약 없음)</summary>
        static string FQ1(long v) => v.ToString("+#,0;-#,0;0");

        /// <summary>금액(원) 포맷: 1원 단위 정확값</summary>
        static string FW(long v) => v.ToString("+#,0;-#,0;0") + "원";

        /// <summary>거래량 포맷 (평균): 12.3만, 1.2억</summary>
        static string FVol(double v)
        {
            if (v <= 0) return "—";
            if (v >= 1e8) return (v / 1e8).ToString("F1") + "억";
            if (v >= 1e4) return (v / 1e4).ToString("F1") + "만";
            return v.ToString("N0");
        }

        /// <summary>업종수급 금액 포맷: +12.3억</summary>
        static string FA(double v) => (v / 1e8).ToString("+#,0.0억;-#,0.0억;0억");

        static Color NQ(long v) => v > 0 ? UP_RED : v < 0 ? DOWN_BLUE : TXT_MUTE;
        static Color SC(double s) => s >= 80 ? GREEN : s >= 50 ? TEAL_D : s >= 30 ? AMBER : CORAL;

        void SetRun(bool v)
        {
            _running = v; _btnRun.Enabled = !v && _codes.Count > 0; _btnStop.Enabled = v; _btnCsv.Enabled = !v;
            if (_chkAutoReanalyze != null) _chkAutoReanalyze.Enabled = !v;
            if (_chkIntradayRefresh != null) _chkIntradayRefresh.Enabled = !v;
            if (v) { _btnRun.Text = "⏳ 분석 중"; _btnRun.Invalidate(); }
            else { _btnRun.Text = "✅ 분석 완료"; _btnRun.Invalidate(); _bar.Value = 0; _bar.Invalidate(); _lblProg.Text = "분석 완료!"; }
        }
        void InvUI(Action a) { if (InvokeRequired) Invoke(a); else a(); }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveWindowBounds();
            _cts?.Cancel(); _consensus?.Dispose(); base.OnFormClosing(e);
        }

        void RestoreWindowBounds()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var lines = File.ReadAllLines(SettingsPath);
                    if (lines.Length >= 5)
                    {
                        int x = int.Parse(lines[0]), y = int.Parse(lines[1]);
                        int w = int.Parse(lines[2]), h = int.Parse(lines[3]);
                        bool max = lines[4] == "1";
                        if (lines.Length >= 6) _autoReanalyzeEnabled = lines[5] == "1";
                        if (lines.Length >= 7) _intradayRefreshEnabled = lines[6] == "1";

                        var screen = Screen.FromPoint(new Point(x, y));
                        var wa = screen.WorkingArea;
                        if (x >= wa.Left - 50 && y >= wa.Top - 50 && x < wa.Right && y < wa.Bottom)
                        {
                            StartPosition = FormStartPosition.Manual;
                            Location = new Point(x, y);
                            Size = new Size(Math.Max(w, MinimumSize.Width), Math.Max(h, MinimumSize.Height));
                            if (max) WindowState = FormWindowState.Maximized;
                            return;
                        }
                    }
                }
            }
            catch { }

            var area = Screen.PrimaryScreen.WorkingArea;
            int dw = Math.Min(1400, area.Width - 40);
            int dh = Math.Min(area.Height - 40, 820);
            Size = new Size(dw, dh);
            StartPosition = FormStartPosition.CenterScreen;
        }

        void SaveWindowBounds()
        {
            try
            {
                var dir = Path.GetDirectoryName(SettingsPath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
                bool max = WindowState == FormWindowState.Maximized;
                File.WriteAllLines(SettingsPath, new[]
                {
                    bounds.X.ToString(), bounds.Y.ToString(),
                    bounds.Width.ToString(), bounds.Height.ToString(),
                    max ? "1" : "0",
                    _autoReanalyzeEnabled ? "1" : "0",
                    _intradayRefreshEnabled ? "1" : "0"
                });
            }
            catch { }
        }

        sealed class CI { public string Idx, Nm; public CI(string i, string n) { Idx = i; Nm = n; } public override string ToString() => Nm; }
    }
}