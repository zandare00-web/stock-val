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

    // 점수 게이지 바 (세부 패널용)
    internal sealed class ScoreGauge : Control
    {
        public string Label    = "";
        public double Value;
        public double Max      = 100;
        public Color  BarColor = Color.FromArgb(0, 200, 255);

        public ScoreGauge()
        {
            Height = 38;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.DoubleBuffer, true);
            BackColor = Color.FromArgb(13, 21, 32); // BG_CARD
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // 레이블
            TextRenderer.DrawText(g, Label, new Font("Segoe UI", 7.5f),
                new Rectangle(0, 0, 90, 18), Color.FromArgb(90, 120, 155),
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter);

            // 분모
            if (Max > 0)
                TextRenderer.DrawText(g, $"/ {Max:F0}",
                    new Font("Segoe UI", 7f),
                    new Rectangle(Width - 112, 0, 56, 18),
                    Color.FromArgb(50, 78, 110),
                    TextFormatFlags.Right | TextFormatFlags.VerticalCenter);

            // 값
            TextRenderer.DrawText(g, Value.ToString("F1"),
                new Font("Consolas", 9.5f, FontStyle.Bold),
                new Rectangle(Width - 54, 0, 54, 18), BarColor,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter);

            // 트랙
            using (var b = new SolidBrush(Color.FromArgb(18, 32, 50)))
                g.FillRectangle(b, 0, 24, Width, 7);

            // 채우기 바
            double pct  = Max > 0 ? Math.Min(1.0, Math.Max(0, Value / Max)) : 0;
            int    fillW = Math.Max(0, (int)(Width * pct));
            if (fillW > 3)
            {
                var gr = new Rectangle(0, 24, fillW + 1, 7);
                using (var grad = new LinearGradientBrush(gr,
                    Color.FromArgb(Math.Min(255, BarColor.R / 2 + 8),
                                   Math.Min(255, BarColor.G / 2 + 8),
                                   Math.Min(255, BarColor.B / 2 + 8)),
                    BarColor, 0f))
                    g.FillRectangle(grad, 0, 24, fillW, 7);

                // 끝 글로우
                using (var b = new SolidBrush(Color.FromArgb(200, BarColor)))
                    g.FillEllipse(b, fillW - 3, 22, 8, 8);
            }
        }
    }

    // KPI 스탯 칩
    internal sealed class StatChip : Control
    {
        public string Label     = "";
        public string ValueText = "—";
        public Color  Accent    = Color.FromArgb(0, 200, 255);

        public StatChip()
        {
            Size = new Size(105, 50);
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.DoubleBuffer, true);
            BackColor = Color.FromArgb(11, 18, 27); // BG_PANEL
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var rc = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var b = new SolidBrush(Color.FromArgb(14, 22, 34)))
                g.FillRectangle(b, rc);
            using (var pen = new Pen(Color.FromArgb(26, 42, 62)))
                g.DrawRectangle(pen, rc);

            // 상단 액센트 라인
            using (var grad = new LinearGradientBrush(
                new Rectangle(1, 0, Width - 2, 2),
                Color.FromArgb(0, Accent), Accent, 0f))
                g.FillRectangle(grad, 1, 0, Width - 2, 2);

            // 큰 값
            TextRenderer.DrawText(g, ValueText,
                new Font("Consolas", 13f, FontStyle.Bold),
                new Rectangle(0, 5, Width, 28), Accent,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

            // 레이블
            TextRenderer.DrawText(g, Label,
                new Font("Segoe UI", 7f),
                new Rectangle(0, 33, Width, 14),
                Color.FromArgb(68, 98, 128),
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    // 다크 둥근 버튼
    internal sealed class DkBtn : Control
    {
        public Color Bg, Fg;
        bool _h, _dn;

        public DkBtn(string text, Color bg, Color fg, int w, int h)
        {
            Text = text; Bg = bg; Fg = fg;
            Size = new Size(w, h);
            Font = new Font("Segoe UI Semibold", 8.5f);
            Cursor = Cursors.Hand;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.DoubleBuffer, true);
            BackColor = Color.FromArgb(11, 18, 27); // BG_PANEL
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var rc = new Rectangle(0, 0, Width - 1, Height - 1);
            Color bg = Enabled ? (_dn ? Adj(Bg, -22) : _h ? Adj(Bg, 18) : Bg)
                               : Color.FromArgb(20, 33, 50);
            Color fg = Enabled ? Fg : Color.FromArgb(48, 72, 100);

            using (var path = RRPath(rc, 4))
            {
                using (var b = new SolidBrush(bg))    g.FillPath(b, path);
                using (var pen = new Pen(Color.FromArgb(55, fg))) g.DrawPath(pen, path);
            }
            TextRenderer.DrawText(g, Text, Font, rc, fg,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        protected override void OnMouseEnter(EventArgs e) { _h  = true;  Invalidate(); }
        protected override void OnMouseLeave(EventArgs e) { _h  = false; Invalidate(); }
        protected override void OnMouseDown(MouseEventArgs e) { _dn = true;  Invalidate(); }
        protected override void OnMouseUp  (MouseEventArgs e) { _dn = false; Invalidate(); }

        static Color Adj(Color c, int d) =>
            Color.FromArgb(Math.Min(255, Math.Max(0, c.R + d)),
                           Math.Min(255, Math.Max(0, c.G + d)),
                           Math.Min(255, Math.Max(0, c.B + d)));

        static GraphicsPath RRPath(Rectangle r, int d)
        {
            var p = new GraphicsPath(); int dd = d * 2;
            p.AddArc(r.X,          r.Y,           dd, dd, 180, 90);
            p.AddArc(r.Right - dd, r.Y,           dd, dd, 270, 90);
            p.AddArc(r.Right - dd, r.Bottom - dd, dd, dd,   0, 90);
            p.AddArc(r.X,          r.Bottom - dd, dd, dd,  90, 90);
            p.CloseFigure(); return p;
        }
    }

    // 얇은 진행 바
    internal sealed class ThinProg : Control
    {
        public int   Value;
        public Color Bar = Color.FromArgb(0, 200, 255);

        public ThinProg()
        {
            Height = 3;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.DoubleBuffer, true);
            BackColor = Color.FromArgb(11, 18, 27); // BG_PANEL
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            using (var b = new SolidBrush(Color.FromArgb(18, 32, 50)))
                e.Graphics.FillRectangle(b, 0, 0, Width, Height);
            int w = (int)(Width * Math.Min(100, Math.Max(0, Value)) / 100.0);
            if (w > 0)
                using (var b = new SolidBrush(Bar))
                    e.Graphics.FillRectangle(b, 0, 0, w, Height);
        }
    }


    // ══════════════════════════════════════════════════════════
    //  MainForm — Obsidian Terminal
    // ══════════════════════════════════════════════════════════
    public partial class MainForm : Form
    {
        // ── 컬러 팔레트 ───────────────────────────────────────
        static readonly Color BG_ROOT  = Color.FromArgb(7,  11, 17);
        static readonly Color BG_PANEL = Color.FromArgb(11, 18, 27);
        static readonly Color BG_CARD  = Color.FromArgb(13, 21, 32);
        static readonly Color BG_ALT   = Color.FromArgb(16, 26, 39);
        static readonly Color BG_SEL   = Color.FromArgb(0,  38, 62);
        static readonly Color BORDER   = Color.FromArgb(26, 42, 62);

        static readonly Color CYAN     = Color.FromArgb(0,  200, 255);
        static readonly Color CYAN_D   = Color.FromArgb(0,  140, 190);
        static readonly Color GOLD     = Color.FromArgb(255, 185,  0);
        static readonly Color GRN      = Color.FromArgb(0,  230, 118);
        static readonly Color RED      = Color.FromArgb(255,  60, 82);
        static readonly Color VIOLET   = Color.FromArgb(140,  90, 255);
        static readonly Color AMBER    = Color.FromArgb(255, 160,  30);

        static readonly Color TXT1     = Color.FromArgb(215, 232, 248);
        static readonly Color TXT2     = Color.FromArgb(100, 130, 162);
        static readonly Color TXT3     = Color.FromArgb(50,  78, 110);

        // ── 상태 ──
        AxKHOpenAPI _ax;
        List<string> _codes = new List<string>();
        List<AnalysisResult> _res = new List<AnalysisResult>();
        List<SectorSupplySummary> _sK = new List<SectorSupplySummary>(),
                                  _sD = new List<SectorSupplySummary>();
        CancellationTokenSource _cts;
        bool _running;

        // ── 컨트롤 ──
        DkBtn    _btnLogin, _btnCsv, _btnRun, _btnStop;
        ComboBox _cbCond;
        Label    _lblLogin, _lblCsv, _lblProg, _lcond;
        ThinProg _bar;
        DataGridView _gStock, _gResult, _gSector;
        Panel    _pDetail;
        StatChip _kpiTotal, _kpiAvg, _kpiGood, _kpiSect;

        public MainForm() { InitializeComponent(); Load += (s, e) => { BuildOcx(); BuildUI(); }; }

        void BuildOcx()
        {
            try
            {
                _ax = new AxKHOpenAPI();
                ((System.ComponentModel.ISupportInitialize)_ax).BeginInit();
                _ax.Visible = false; _ax.Width = 1; _ax.Height = 1;
                Controls.Add(_ax);
                ((System.ComponentModel.ISupportInitialize)_ax).EndInit();
            }
            catch (Exception ex) { MessageBox.Show("키움 OCX 오류:\n" + ex.Message); }
        }


        // ═══════════════ UI 빌드 ════════════════════════════

        void BuildUI()
        {
            Text = "Stock Analyzer";
            Size = new Size(1500, 920);
            MinimumSize = new Size(1100, 720);
            BackColor = BG_ROOT; ForeColor = TXT1;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9f);
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.DoubleBuffer, true);

            // ─── 헤더 (68px) ─────────────────────────────────
            var hdr = new Panel { Dock = DockStyle.Top, Height = 68, BackColor = BG_PANEL };
            hdr.Paint  += PaintHeader;
            hdr.Resize += (s, e) => LayoutHeader(hdr);

            _lblLogin = new Label { Text = "● 미연결", ForeColor = RED,
                Font = new Font("Segoe UI", 8f), BackColor = BG_PANEL,
                TextAlign = ContentAlignment.MiddleRight };
            _btnLogin = new DkBtn("연결", Color.FromArgb(0, 48, 74), CYAN, 60, 28);
            _btnLogin.Click += BtnLogin_Click;

            _btnCsv = new DkBtn("📂 CSV", Color.FromArgb(18, 30, 46), TXT2, 70, 28);
            _btnCsv.Click += BtnCsv_Click;
            _lblCsv = new Label { Text = "파일 없음", ForeColor = TXT3,
                Font = new Font("Segoe UI", 7.5f), BackColor = BG_PANEL,
                TextAlign = ContentAlignment.MiddleLeft };

            _lcond = new Label { Text = "조건검색", ForeColor = TXT3,
                Font = new Font("Segoe UI", 7.5f), BackColor = BG_PANEL,
                TextAlign = ContentAlignment.MiddleLeft, Width = 46 };
            _cbCond = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(16, 28, 44), ForeColor = TXT1,
                FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5f) };
            _cbCond.SelectedIndexChanged += CbCond_Changed;

            _btnRun  = new DkBtn("▶  분석 시작", Color.FromArgb(0,  58, 92), CYAN, 104, 28) { Enabled = false };
            _btnStop = new DkBtn("■  중지",      Color.FromArgb(72, 18, 28), RED,   68, 28) { Enabled = false };
            _btnRun.Click  += BtnRun_Click;
            _btnStop.Click += (s, e) => _cts?.Cancel();

            _bar     = new ThinProg { Width = 110, Bar = CYAN };
            _lblProg = new Label { Text = "대기 중", ForeColor = TXT3,
                Font = new Font("Segoe UI", 7.5f), BackColor = BG_PANEL,
                TextAlign = ContentAlignment.MiddleLeft };

            _kpiTotal = new StatChip { Label = "분석 종목",  ValueText = "0",  Accent = CYAN };
            _kpiAvg   = new StatChip { Label = "평균 총점",  ValueText = "—",  Accent = GOLD };
            _kpiGood  = new StatChip { Label = "수급 양호",  ValueText = "0",  Accent = GRN };
            _kpiSect  = new StatChip { Label = "업종 수",    ValueText = "0",  Accent = VIOLET };

            hdr.Controls.AddRange(new Control[]
            {
                _lblLogin, _btnLogin, _btnCsv, _lblCsv, _lcond, _cbCond,
                _btnRun, _btnStop, _bar, _lblProg,
                _kpiTotal, _kpiAvg, _kpiGood, _kpiSect
            });

            var sep = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = BORDER };

            // ─── 본문 3열 ─────────────────────────────────────
            var body = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1,
                BackColor = BG_ROOT, Padding = new Padding(8, 8, 8, 8)
            };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 264));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 320));

            body.Controls.Add(BuildLeft(),   0, 0);
            body.Controls.Add(BuildCenter(), 1, 0);
            body.Controls.Add(BuildRight(),  2, 0);

            Controls.Add(body);
            Controls.Add(sep);
            Controls.Add(hdr);
        }


        // ─── 헤더 페인트 ──────────────────────────────────────

        void PaintHeader(object s, PaintEventArgs e)
        {
            var g = e.Graphics; var p = (Panel)s;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // 좌측 수직 시안 라인
            using (var grad = new LinearGradientBrush(
                new Rectangle(0, 0, 3, p.Height),
                CYAN, Color.FromArgb(0, CYAN), 90f))
                g.FillRectangle(grad, 0, 0, 3, p.Height);

            // 로고
            TextRenderer.DrawText(g, "STA",
                new Font("Consolas", 16f, FontStyle.Bold),
                new Rectangle(16, 8, 58, 26), CYAN, TextFormatFlags.Left);
            TextRenderer.DrawText(g, "◆",
                new Font("Segoe UI", 10f),
                new Rectangle(64, 8, 20, 26),
                Color.FromArgb(0, 118, 168),
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            TextRenderer.DrawText(g, "STOCK  ANALYZER",
                new Font("Consolas", 7f),
                new Rectangle(16, 36, 155, 16), TXT3, TextFormatFlags.Left);

            // 하단 보더
            using (var pen = new Pen(BORDER))
                g.DrawLine(pen, 0, p.Height - 1, p.Width, p.Height - 1);

            // KPI 구분선
            if (_kpiTotal != null)
            {
                int sx = _kpiTotal.Left - 12;
                if (sx > 0)
                    using (var pen = new Pen(BORDER))
                        g.DrawLine(pen, sx, 8, sx, p.Height - 8);
            }
        }


        // ─── 헤더 레이아웃 ────────────────────────────────────

        void LayoutHeader(Panel p)
        {
            int h = p.Height;

            int rx = p.Width - 10;
            rx -= 60;  _btnLogin.SetBounds(rx, (h - 28) / 2, 60, 28);
            rx -= 128; _lblLogin.SetBounds(rx, 0, 124, h);

            rx -= 14;
            rx -= 105; _kpiSect.SetBounds(rx, (h - 50) / 2, 105, 50);
            rx -= 109; _kpiGood.SetBounds(rx, (h - 50) / 2, 105, 50);
            rx -= 109; _kpiAvg.SetBounds(rx,  (h - 50) / 2, 105, 50);
            rx -= 109; _kpiTotal.SetBounds(rx, (h - 50) / 2, 105, 50);
            p.Invalidate();

            int lx = 162;
            _btnCsv.SetBounds(lx,  (h - 28) / 2, 70, 28);  lx += 78;
            _lblCsv.SetBounds(lx,  0, 80, h);               lx += 84;
            _lcond.SetBounds(lx,   0, 48, h);               lx += 50;
            _cbCond.SetBounds(lx,  (h - 28) / 2, 160, 28); lx += 168;
            _btnRun.SetBounds(lx,  (h - 28) / 2, 104, 28); lx += 112;
            _btnStop.SetBounds(lx, (h - 28) / 2, 68, 28);  lx += 76;
            _bar.SetBounds(lx, h / 2 + 3, 100, 3);         lx += 108;
            _lblProg.SetBounds(lx, 0, 170, h);
        }


        // ─── 좌측: Watchlist ──────────────────────────────────

        Control BuildLeft()
        {
            var wrap = new Panel { Dock = DockStyle.Fill, BackColor = BG_ROOT,
                                   Padding = new Padding(0, 0, 5, 0) };
            var card = MkCard("WATCHLIST", "종목 리스트");

            _gStock = DkGrid(
                ("종목명", "Name",   0,  false),
                ("코드",   "Code",   64, false),
                ("시장",   "Market", 52, false)
            );
            _gStock.SelectionChanged += GSel;
            card.Controls.Add(_gStock);
            wrap.Controls.Add(card);
            return wrap;
        }


        // ─── 중앙: Analysis Results ───────────────────────────

        Control BuildCenter()
        {
            var card = MkCard("ANALYSIS", "분석 결과");
            _gResult = DkGrid(
                ("#",       "Rank",             32,  true),
                ("종목명",  "Name",              85, false),
                ("총점",    "TotalScore",        52,  true),
                ("수급",    "StockSupplyScore",  56,  true),
                ("외국인",  "ForeignNet5D",      70,  true),
                ("기관",    "InstNet5D",         70,  true),
                ("추세",    "SupplyTrendStr",    48, false),
                ("업종",    "SectorName",        68, false),
                ("업종수급","SectorSupplyScore", 58,  true)
            );
            _gResult.SelectionChanged += GRSel;
            _gResult.CellFormatting   += GRFmt;
            card.Controls.Add(_gResult);
            return card;
        }


        // ─── 우측: Detail + Sector ────────────────────────────

        Control BuildRight()
        {
            var wrap = new Panel { Dock = DockStyle.Fill, BackColor = BG_ROOT,
                                   Padding = new Padding(5, 0, 0, 0) };
            var split = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = BG_ROOT
            };
            split.RowStyles.Add(new RowStyle(SizeType.Percent, 57));
            split.RowStyles.Add(new RowStyle(SizeType.Percent, 43));

            var detCard = MkCard("DETAIL", "종목 세부 정보");
            _pDetail = new Panel { Dock = DockStyle.Fill, BackColor = BG_CARD, AutoScroll = true };
            ShowDetail(null);
            detCard.Controls.Add(_pDetail);

            var secCard = MkCard("SECTOR", "업종 수급 현황");
            _gSector = DkGrid(
                ("업종명", "SectorName",    0,  false),
                ("시장",   "Market",        44, false),
                ("외국인", "ForeignNet5DB", 62,  true),
                ("기관",   "InstNet5DB",    62,  true),
                ("합산",   "TotalNet5DB",   62,  true)
            );
            secCard.Controls.Add(_gSector);

            split.Controls.Add(detCard, 0, 0);
            split.Controls.Add(secCard, 0, 1);
            wrap.Controls.Add(split);
            return wrap;
        }


        // ─── 카드 팩토리 ─────────────────────────────────────

        static Panel MkCard(string tag, string title)
        {
            var card = new Panel { Dock = DockStyle.Fill, BackColor = BG_CARD };
            card.Paint += (s, e) =>
            {
                using (var pen = new Pen(BORDER))
                    e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
            };

            var chdr = new Panel { Dock = DockStyle.Top, Height = 32, BackColor = BG_PANEL };
            chdr.Paint += (s, e) =>
            {
                var g = e.Graphics; var h = (Panel)s;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // 하단 보더
                using (var pen = new Pen(BORDER))
                    g.DrawLine(pen, 0, h.Height - 1, h.Width, h.Height - 1);

                // 태그 배지
                var tagFont = new Font("Consolas", 7.5f);
                var sz      = TextRenderer.MeasureText(tag, tagFont);
                int tw = sz.Width + 18, th = 18;
                int tx = 10,            ty = (h.Height - th) / 2;

                using (var b = new SolidBrush(Color.FromArgb(0, 46, 74)))
                    g.FillRectangle(b, tx, ty, tw, th);
                using (var pen = new Pen(Color.FromArgb(0, 88, 132)))
                    g.DrawRectangle(pen, tx, ty, tw, th);
                TextRenderer.DrawText(g, tag, tagFont,
                    new Rectangle(tx, ty, tw, th), CYAN,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                tagFont.Dispose();

                // 제목
                TextRenderer.DrawText(g, title, new Font("Segoe UI", 8f),
                    new Rectangle(tx + tw + 10, 0, h.Width - tx - tw - 60, h.Height),
                    TXT2,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
            };

            card.Controls.Add(chdr);
            return card;
        }


        // ─── 그리드 팩토리 ────────────────────────────────────

        static DataGridView DkGrid(params (string h, string n, int w, bool r)[] cols)
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor     = BG_CARD,
                BorderStyle         = BorderStyle.None,
                GridColor           = BORDER,
                CellBorderStyle     = DataGridViewCellBorderStyle.SingleHorizontal,
                RowHeadersVisible   = false,
                AllowUserToAddRows  = false,
                AllowUserToDeleteRows = false,
                ReadOnly            = true,
                SelectionMode       = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeight = 26,
                RowTemplate         = { Height = 25 },
                EnableHeadersVisualStyles = false,
                ScrollBars          = ScrollBars.Vertical,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor          = BG_CARD,
                    ForeColor          = TXT1,
                    SelectionBackColor = BG_SEL,
                    SelectionForeColor = CYAN,
                    Font               = new Font("Segoe UI", 8f),
                    Padding            = new Padding(5, 0, 5, 0)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor          = BG_PANEL,
                    ForeColor          = TXT3,
                    Font               = new Font("Segoe UI Semibold", 7.5f),
                    SelectionBackColor = BG_PANEL,
                    Padding            = new Padding(5, 0, 5, 0)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor          = BG_ALT,
                    ForeColor          = TXT1,
                    SelectionBackColor = BG_SEL,
                    SelectionForeColor = CYAN
                }
            };

            foreach (var (h, n, w, ra) in cols)
                g.Columns.Add(new DataGridViewTextBoxColumn
                {
                    HeaderText   = h,
                    Name         = n,
                    MinimumWidth = w > 0 ? w : 60,
                    FillWeight   = w == 0 ? 200f : 100f,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = ra
                            ? DataGridViewContentAlignment.MiddleRight
                            : DataGridViewContentAlignment.MiddleLeft
                    }
                });

            // 히든 Code2 (분석 결과 그리드 전용)
            if (cols.Any(c => c.n == "Name") && cols.Any(c => c.n == "TotalScore"))
                g.Columns.Add(new DataGridViewTextBoxColumn { Name = "Code2", Visible = false });

            return g;
        }


        // ═══════════════ 이벤트 ═════════════════════════════

        async void BtnLogin_Click(object s, EventArgs e)
        {
            if (_ax == null) return;
            _btnLogin.Enabled = false;
            _lblLogin.Text = "● 연결 중..."; _lblLogin.ForeColor = AMBER;
            try
            {
                using (var cl = new KiwoomClient(_ax))
                {
                    if (await cl.LoginAsync())
                    {
                        var nm = _ax.GetLoginInfo("USER_NAME");
                        var sv = _ax.GetLoginInfo("GetServerGubun") == "1" ? "모의" : "실";
                        _lblLogin.Text = $"● {nm} ({sv})";
                        _lblLogin.ForeColor = GRN;
                        _btnLogin.Bg = Color.FromArgb(0, 46, 28);
                        _btnLogin.Fg = GRN;
                        _btnLogin.Text = "연결됨";
                        _btnLogin.Invalidate();
                        await LoadConds(cl);
                        UpdRun();
                    }
                }
            }
            catch (Exception ex)
            {
                _lblLogin.Text = "● 실패"; _lblLogin.ForeColor = RED;
                _btnLogin.Enabled = true; _btnLogin.Invalidate();
                MessageBox.Show("로그인 실패: " + ex.Message);
            }
        }

        async Task LoadConds(KiwoomClient cl)
        {
            try
            {
                var ls = await cl.GetConditionListAsync();
                _cbCond.Items.Clear();
                _cbCond.Items.Add(new CI("", "— 조건 선택 —"));
                foreach (var (i, n) in ls) _cbCond.Items.Add(new CI(i, n));
                if (_cbCond.Items.Count > 0) _cbCond.SelectedIndex = 0;
            }
            catch { }
        }

        void BtnCsv_Click(object s, EventArgs e)
        {
            using (var d = new OpenFileDialog { Filter = "CSV|*.csv|All|*.*" })
            {
                if (d.ShowDialog() != DialogResult.OK) return;
                try
                {
                    LoadCds(CsvCodeExtractor.Extract(d.FileName));
                    _lblCsv.Text = Path.GetFileName(d.FileName);
                    _lblCsv.ForeColor = CYAN_D;
                }
                catch (Exception ex) { MessageBox.Show("CSV 오류: " + ex.Message); }
            }
        }

        async void CbCond_Changed(object s, EventArgs e)
        {
            if (_cbCond.SelectedItem is CI ci && ci.Idx != "")
            {
                if (_ax == null) return;
                try
                {
                    using (var c = new KiwoomClient(_ax))
                        LoadCds(await c.GetConditionCodesAsync(ci.Idx, ci.Nm));
                }
                catch (Exception ex) { MessageBox.Show("조건검색 실패: " + ex.Message); }
            }
        }

        void LoadCds(List<string> c)
        {
            _codes = c;
            _gStock.Rows.Clear();
            foreach (var x in c) _gStock.Rows.Add("—", x, "");
            UpdRun();
            _kpiTotal.ValueText = c.Count.ToString();
            _kpiTotal.Invalidate();
        }

        void UpdRun() => _btnRun.Enabled = _ax?.GetConnectState() == 1 && _codes.Count > 0;

        async void BtnRun_Click(object s, EventArgs e)
        {
            if (_running || _codes.Count == 0) return;
            SetRun(true);
            _res.Clear(); _gResult.Rows.Clear();
            _cts = new CancellationTokenSource();
            var eng = new AnalysisEngine(_ax);
            eng.Progress += (cur, tot, nm) => InvUI(() =>
            {
                _bar.Value = (int)((double)cur / tot * 100);
                _bar.Invalidate();
                _lblProg.Text = $"{nm}  {cur}/{tot}";
            });
            eng.Log += m => System.Diagnostics.Debug.WriteLine(m);
            try
            {
                var (r, sk, sd) = await eng.RunAsync(_codes, ScoreConfig.Instance, _cts.Token);
                _res = r; _sK = sk; _sD = sd;
                FillResult(); FillSector(); FillStock(); UpdateKpi();
            }
            catch (OperationCanceledException) { _lblProg.Text = "중지됨"; }
            catch (Exception ex) { MessageBox.Show("분석 오류: " + ex.Message); }
            finally { SetRun(false); }
        }

        void GSel(object s, EventArgs e)
        {
            if (_gStock.SelectedRows.Count == 0) return;
            var cd = _gStock.SelectedRows[0].Cells["Code"].Value?.ToString();
            ShowDetail(_res.FirstOrDefault(r => r.Code == cd));
            for (int i = 0; i < _gResult.Rows.Count; i++)
                if (_gResult.Rows[i].Cells["Code2"]?.Value?.ToString() == cd)
                { _gResult.Rows[i].Selected = true; break; }
        }

        void GRSel(object s, EventArgs e)
        {
            if (_gResult.SelectedRows.Count == 0) return;
            var cd = _gResult.SelectedRows[0].Cells["Code2"]?.Value?.ToString();
            ShowDetail(_res.FirstOrDefault(r => r.Code == cd));
        }

        void GRFmt(object s, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _res.Count) return;
            var r   = _res[e.RowIndex];
            var col = _gResult.Columns[e.ColumnIndex].Name;

            switch (col)
            {
                case "TotalScore":
                    e.CellStyle.ForeColor = ScoreClr(r.TotalScore);
                    e.CellStyle.Font = new Font("Consolas", 8.5f, FontStyle.Bold);
                    break;
                case "StockSupplyScore":
                    if (double.TryParse(e.Value?.ToString(), out double sv))
                        e.CellStyle.ForeColor = sv >= 40 ? CYAN : sv >= 20 ? GOLD : TXT2;
                    break;
                case "SupplyTrendStr":
                    e.CellStyle.ForeColor =
                        r.SupplyTrend == SupplyTrend.상승 || r.SupplyTrend == SupplyTrend.상승반전 ? GRN
                        : r.SupplyTrend == SupplyTrend.하락반전 ? AMBER
                        : r.SupplyTrend == SupplyTrend.하락     ? RED
                        : TXT3;
                    break;
                case "ForeignNet5D":
                case "InstNet5D":
                    if (e.Value?.ToString() is string sv2)
                        e.CellStyle.ForeColor = sv2.StartsWith("+") ? GRN
                                              : sv2.StartsWith("-") ? RED
                                              : TXT2;
                    break;
            }
        }


        // ═══════════════ 데이터 갱신 ════════════════════════

        void FillResult()
        {
            _gResult.Rows.Clear();
            for (int i = 0; i < _res.Count; i++)
            {
                var r = _res[i];
                _gResult.Rows.Add(
                    i + 1, r.Name,
                    r.TotalScore.ToString("F1"),
                    r.StockSupplyScore.ToString("F1"),
                    FN(r.ForeignNet5D), FN(r.InstNet5D),
                    r.SupplyTrend.ToString(),
                    r.SectorName,
                    r.SectorSupplyScore.ToString("F1"),
                    r.Code);
            }
        }

        void FillSector()
        {
            _gSector.Rows.Clear();
            foreach (var x in _sK.Concat(_sD).OrderByDescending(x => x.TotalNet5D))
                _gSector.Rows.Add(x.SectorName, x.Market, FA(x.ForeignNet5D), FA(x.InstNet5D), FA(x.TotalNet5D));
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
            _kpiTotal.ValueText = _res.Count.ToString(); _kpiTotal.Invalidate();
            if (_res.Count > 0) { _kpiAvg.ValueText = _res.Average(r => r.TotalScore).ToString("F1"); _kpiAvg.Invalidate(); }
            _kpiGood.ValueText = _res.Count(r => r.StockSupplyScore >= 50).ToString(); _kpiGood.Invalidate();
            int sects = _res.Select(r => r.SectorName).Where(n => !string.IsNullOrEmpty(n)).Distinct().Count();
            _kpiSect.ValueText = sects.ToString(); _kpiSect.Invalidate();
        }


        // ═══════════════ 세부 패널 ══════════════════════════

        void ShowDetail(AnalysisResult r)
        {
            _pDetail.Controls.Clear();

            if (r == null)
            {
                _pDetail.Controls.Add(new Label
                {
                    Text = "← 종목을 선택하세요",
                    ForeColor = TXT3,
                    Font = new Font("Segoe UI", 8.5f),
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = BG_CARD
                });
                return;
            }

            var fl = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                BackColor = BG_CARD,
                Padding = new Padding(14, 12, 14, 12)
            };

            // 종목 헤더
            fl.Controls.Add(DLbl($"{r.Name}", new Font("Segoe UI Semibold", 11f), TXT1));
            fl.Controls.Add(DLbl($"{r.Code}  ·  {r.Market}  ·  {r.SectorName}", new Font("Segoe UI", 7.5f), TXT3));
            fl.Controls.Add(DLbl($"현재가  {r.CurrentPrice:N0}원", new Font("Consolas", 9f), CYAN_D));
            fl.Controls.Add(DSep(14));

            // 점수 게이지
            fl.Controls.Add(DSection("SCORE"));
            var cfg = ScoreConfig.Instance;
            fl.Controls.Add(MkGauge("총점",     r.TotalScore,        cfg.TotalMaxScore,         ScoreClr(r.TotalScore)));
            fl.Controls.Add(MkGauge("기업가치", r.ValueScore,        cfg.TotalValueScore,        GOLD));
            fl.Controls.Add(MkGauge("종목수급", r.StockSupplyScore,  cfg.TotalStockSupplyScore,  CYAN));
            fl.Controls.Add(MkGauge("업종수급", r.SectorSupplyScore, cfg.TotalSectorSupplyScore, VIOLET));
            fl.Controls.Add(DSep(12));

            // 펀더멘털
            fl.Controls.Add(DSection("FUNDAMENTAL"));
            fl.Controls.Add(DRow("PER",
                r.Per.HasValue  ? r.Per.Value.ToString("F2")  : "—",
                r.SectorAvgPer.HasValue ? $"업종 {r.SectorAvgPer.Value:F2}" : ""));
            fl.Controls.Add(DRow("PBR",
                r.Pbr.HasValue  ? r.Pbr.Value.ToString("F2")  : "—",
                r.SectorAvgPbr.HasValue ? $"업종 {r.SectorAvgPbr.Value:F2}" : ""));
            fl.Controls.Add(DRow("ROE",
                r.Roe.HasValue  ? r.Roe.Value.ToString("F1") + "%" : "—", ""));
            fl.Controls.Add(DSep(12));

            // 수급
            fl.Controls.Add(DSection("SUPPLY / DEMAND"));
            fl.Controls.Add(DRow("외국인 1D",  FN(r.ForeignNetD1),  "", NC(r.ForeignNetD1)));
            fl.Controls.Add(DRow("외국인 5D",  FN(r.ForeignNet5D),  "", NC(r.ForeignNet5D)));
            fl.Controls.Add(DRow("외국인 20D", FN(r.ForeignNet20D), "", NC(r.ForeignNet20D)));
            fl.Controls.Add(DRow("기관 1D",    FN(r.InstNetD1),     "", NC(r.InstNetD1)));
            fl.Controls.Add(DRow("기관 5D",    FN(r.InstNet5D),     "", NC(r.InstNet5D)));
            fl.Controls.Add(DRow("기관 20D",   FN(r.InstNet20D),    "", NC(r.InstNet20D)));
            fl.Controls.Add(DSep(12));

            // 추세
            fl.Controls.Add(DSection("TREND"));
            fl.Controls.Add(DRow("회전율 추세",
                r.TurnoverRate.ToString("+0.0;-0.0") + "%", "",
                r.TurnoverRate > 0 ? GRN : r.TurnoverRate < 0 ? RED : TXT2));
            fl.Controls.Add(DRow("수급 추세", r.SupplyTrend.ToString(), "",
                r.SupplyTrend == SupplyTrend.상승 || r.SupplyTrend == SupplyTrend.상승반전 ? GRN
                : r.SupplyTrend == SupplyTrend.하락 ? RED : AMBER));

            _pDetail.Controls.Add(fl);
        }


        // ─── 세부 패널 헬퍼 ──────────────────────────────────

        static ScoreGauge MkGauge(string lbl, double val, double max, Color clr)
            => new ScoreGauge
            {
                Label = lbl, Value = val, Max = max, BarColor = clr,
                Width = 278, Margin = new Padding(0, 2, 0, 2)
            };

        static Label DLbl(string t, Font f, Color c)
            => new Label
            {
                Text = t, AutoSize = false, Width = 278,
                Height = f.Height + 6, Font = f, ForeColor = c,
                BackColor = BG_CARD, Margin = new Padding(0, 1, 0, 1)
            };

        static Label DSection(string t)
            => new Label
            {
                Text = t, AutoSize = false, Width = 278, Height = 16,
                Font = new Font("Consolas", 6.8f), ForeColor = TXT3,
                BackColor = BG_CARD, Margin = new Padding(0, 2, 0, 3)
            };

        static Panel DSep(int margin = 6)
            => new Panel
            {
                Width = 278, Height = 1, BackColor = BORDER,
                Margin = new Padding(0, margin, 0, margin)
            };

        static Panel DRow(string lbl, string val, string sub = "", Color? vc = null)
        {
            var p = new Panel { Width = 278, Height = 22, BackColor = BG_CARD };
            p.Controls.Add(new Label
            {
                Text = lbl, Width = 80, ForeColor = TXT2,
                Font = new Font("Segoe UI", 7.5f),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Left, BackColor = BG_CARD
            });
            if (!string.IsNullOrEmpty(sub))
                p.Controls.Add(new Label
                {
                    Text = sub, Width = 70, ForeColor = TXT3,
                    Font = new Font("Segoe UI", 7f),
                    TextAlign = ContentAlignment.MiddleRight,
                    Dock = DockStyle.Right, BackColor = BG_CARD
                });
            p.Controls.Add(new Label
            {
                Text = val, ForeColor = vc ?? TXT1,
                Font = new Font("Consolas", 8.5f),
                TextAlign = ContentAlignment.MiddleRight,
                Dock = DockStyle.Fill, BackColor = BG_CARD
            });
            return p;
        }


        // ─── 공통 유틸 ───────────────────────────────────────

        static string FN(long v)   => v.ToString("+#,0;-#,0;0");
        static string FA(double v) => (v / 1e8).ToString("+#,0.0억;-#,0.0억;0억");
        static Color  NC(long v)   => v >= 0 ? GRN : RED;
        static Color  ScoreClr(double s) => s >= 80 ? GRN : s >= 50 ? CYAN : s >= 30 ? GOLD : RED;

        void SetRun(bool v)
        {
            _running = v;
            _btnRun.Enabled  = !v && _codes.Count > 0;
            _btnStop.Enabled = v;
            _btnCsv.Enabled  = !v;
            if (!v) { _bar.Value = 0; _bar.Invalidate(); _lblProg.Text = "완료"; }
        }

        void InvUI(Action a) { if (InvokeRequired) Invoke(a); else a(); }

        protected override void OnFormClosing(FormClosingEventArgs e)
        { _cts?.Cancel(); base.OnFormClosing(e); }

        sealed class CI
        {
            public string Idx, Nm;
            public CI(string i, string n) { Idx = i; Nm = n; }
            public override string ToString() => Nm;
        }
    }
}
