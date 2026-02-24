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
        public KpiCard() { BackColor = Color.Transparent; Size = new Size(180, 80); SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var p = RR(new Rectangle(0, 0, Width - 1, Height - 1), 10))
            { using (var b = new SolidBrush(Color.White)) g.FillPath(b, p); using (var pen = new Pen(Color.FromArgb(226, 232, 240))) g.DrawPath(pen, p); }
            TextRenderer.DrawText(g, Title, new Font("Segoe UI", 8.5f, FontStyle.Bold), new Rectangle(16, 12, Width - 32, 16), Color.FromArgb(100, 116, 139), TextFormatFlags.Left);
            TextRenderer.DrawText(g, Val, new Font("Segoe UI", 18f, FontStyle.Bold), new Rectangle(14, 30, Width - 28, 30), ValColor, TextFormatFlags.Left);
            if (!string.IsNullOrEmpty(Sub))
                TextRenderer.DrawText(g, Sub, new Font("Segoe UI", 7.5f), new Rectangle(16, 60, Width - 32, 14), Color.FromArgb(148, 163, 184), TextFormatFlags.Left);
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

        // ── 상태 ──
        AxKHOpenAPI _ax;
        List<string> _codes = new List<string>();
        List<AnalysisResult> _res = new List<AnalysisResult>();
        List<SectorSupplySummary> _sK = new List<SectorSupplySummary>(), _sD = new List<SectorSupplySummary>();
        CancellationTokenSource _cts; bool _running;

        // ── 컨트롤 ──
        RndBtn _btnLogin, _btnCsv, _btnRun, _btnStop, _btnSet;
        ComboBox _cbCond; Label _lblLogin, _lblCsv, _lblProg;
        SlimBar _bar;
        DataGridView _gStock, _gResult, _gSector;
        Panel _pDetail;
        KpiCard _kpiTotal, _kpiValue, _kpiSupply, _kpiSector;

        public MainForm() { InitializeComponent(); Load += (s, e) => { BuildOcx(); BuildUI(); }; }

        void BuildOcx()
        {
            try { _ax = new AxKHOpenAPI(); ((System.ComponentModel.ISupportInitialize)_ax).BeginInit(); _ax.Visible = false; _ax.Width = 1; _ax.Height = 1; Controls.Add(_ax); ((System.ComponentModel.ISupportInitialize)_ax).EndInit(); }
            catch (Exception ex) { MessageBox.Show("키움 OCX 오류:\n" + ex.Message); }
        }

        // ═══════════════ UI 빌드 ═══════════════════════════════

        void BuildUI()
        {
            Text = "Stock Analyzer"; Size = new Size(1400, 900); MinimumSize = new Size(1300, 700);
            BackColor = MAIN_BG; ForeColor = TXT_MAIN; StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9.5f);
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);

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

            // 2. 통합 상단바 (헤더 1줄 통합)
            var topBar = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = MAIN_BG };
            topBar.Paint += (s, e) => { using (var p = new Pen(CARD_BRD)) e.Graphics.DrawLine(p, 24, topBar.Height - 1, topBar.Width - 24, topBar.Height - 1); };

            var title = new Label { Text = "📊 Stock Analyzer", Font = new Font("Segoe UI", 15f, FontStyle.Bold), ForeColor = HDR_TXT, BackColor = MAIN_BG, TextAlign = ContentAlignment.MiddleLeft };

            _btnCsv = new RndBtn("📂 CSV", Color.White, TXT_MAIN, 90, 32) { Bdr = CARD_BRD, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold) };
            _btnCsv.Click += BtnCsv_Click;
            _lblCsv = new Label { Text = "선택된 파일 없음", ForeColor = TXT_MUTE, Font = new Font("Segoe UI", 8.5f), BackColor = MAIN_BG, TextAlign = ContentAlignment.MiddleLeft };

            var lcond = new Label { Text = "조건검색", ForeColor = TXT_SEC, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), BackColor = MAIN_BG, TextAlign = ContentAlignment.MiddleRight };
            _cbCond = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.White, ForeColor = TXT_MAIN, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f) };
            _cbCond.SelectedIndexChanged += CbCond_Changed;

            _btnRun = new RndBtn("▶ 분석 시작", TEAL, Color.White, 100, 32) { Enabled = false };
            _btnRun.Click += BtnRun_Click;
            _btnStop = new RndBtn("■ 중지", CORAL, Color.White, 76, 32) { Enabled = false };
            _btnStop.Click += (s, e) => _cts?.Cancel();

            _btnSet = new RndBtn("⚙ 설정", Color.White, TXT_MAIN, 70, 32) { Bdr = CARD_BRD };
            _btnSet.Click += (s, e) => { using (var f = new SettingsForm()) f.ShowDialog(this); };

            _lblLogin = new Label { ForeColor = Color.FromArgb(244, 63, 94), Font = new Font("Segoe UI", 9f, FontStyle.Bold), BackColor = MAIN_BG, TextAlign = ContentAlignment.MiddleRight, Text = "● 미연결" };
            _btnLogin = new RndBtn("로그인", TEAL, Color.White, 76, 32);
            _btnLogin.Click += BtnLogin_Click;

            topBar.Controls.AddRange(new Control[] { title, _btnCsv, _lblCsv, lcond, _cbCond, _btnRun, _btnStop, _btnSet, _lblLogin, _btnLogin });

            // 상단바 1줄 레이아웃 및 콤보박스 수직 정렬(y+4) 로직
            topBar.Resize += (s, e) => {
                int y = 16; // 64px 높이에서 32px 버튼을 수직 중앙 정렬
                title.SetBounds(24, 0, 200, topBar.Height);

                int x = 224;
                _btnCsv.Location = new Point(x, y); x += _btnCsv.Width + 8;
                _lblCsv.SetBounds(x, y, 110, 32); x += 110;

                lcond.SetBounds(x, y, 60, 32); x += 64;
                // ★ 콤보박스가 위로 치우치지 않도록 +4 픽셀 내림 조정
                _cbCond.SetBounds(x, y + 4, 160, 26); x += 172;

                _btnRun.Location = new Point(x, y); x += _btnRun.Width + 8;
                _btnStop.Location = new Point(x, y); x += _btnStop.Width + 12;
                _btnSet.Location = new Point(x, y);

                _btnLogin.Location = new Point(topBar.Width - 100, y);
                _lblLogin.SetBounds(topBar.Width - 100 - 160, 0, 150, topBar.Height);
            };

            // 3. 본문 
            var body = new Panel { Dock = DockStyle.Fill, BackColor = MAIN_BG, Padding = new Padding(24, 12, 24, 24) };

            var kpiRow = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 96, BackColor = MAIN_BG, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(0, 0, 0, 16) };
            _kpiTotal = new KpiCard { Title = "TOTAL STOCKS", Val = "0", Sub = "분석 대기중" };
            _kpiValue = new KpiCard { Title = "AVG SCORE", Val = "—", Sub = "최대 125점" };
            _kpiSupply = new KpiCard { Title = "STRONG SUPPLY", Val = "0", Sub = "수급 50점 이상", ValColor = TEAL_D };
            _kpiSector = new KpiCard { Title = "SECTORS", Val = "0", Sub = "분석된 업종 수" };
            kpiRow.Controls.AddRange(new Control[] { _kpiTotal, Sp(16), _kpiValue, Sp(16), _kpiSupply, Sp(16), _kpiSector });

            var grid3 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, BackColor = MAIN_BG, Padding = new Padding(0) };
            grid3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));
            grid3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
            grid3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            grid3.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            grid3.Controls.Add(WCard("WATCHLIST", BuildP1()), 0, 0);
            grid3.Controls.Add(WCard("ANALYSIS RESULT", BuildP2()), 1, 0);
            grid3.Controls.Add(BuildP3(), 2, 0);

            body.Controls.Add(grid3);
            body.Controls.Add(kpiRow);

            // 도킹 (역순 추가로 Z-Order 충돌 완전 방어)
            Controls.Add(body);
            Controls.Add(topBar);
            Controls.Add(statusBar);
        }

        // ── 패널 빌더 ──

        Control BuildP1()
        {
            _gStock = MkGrid(("종목명", "Name", 110, false), ("코드", "Code", 60, false), ("시장", "Market", 45, false));
            _gStock.SelectionChanged += GSel;
            return _gStock;
        }

        Control BuildP2()
        {
            _gResult = MkGrid(
                ("순위", "Rank", 40, true), ("종목", "Name", 85, false), ("총점", "TotalScore", 50, true),
                ("종목수급", "StockSupplyScore", 60, true), ("외국인", "ForeignNet5D", 65, true),
                ("기관", "InstNet5D", 65, true), ("추세", "SupplyTrendStr", 50, false),
                ("업종", "SectorName", 65, false), ("업종수급", "SectorSupplyScore", 60, true));
            _gResult.SelectionChanged += GRSel;
            _gResult.CellFormatting += GRFmt;
            return _gResult;
        }

        Control BuildP3()
        {
            var outer = new Panel { Dock = DockStyle.Fill, BackColor = MAIN_BG, Padding = new Padding(12, 0, 0, 0) };
            var sp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = MAIN_BG };
            sp.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            sp.RowStyles.Add(new RowStyle(SizeType.Percent, 60));

            _gSector = MkGrid(("업종", "SectorName", 80, false), ("시장", "Market", 40, false),
                ("외국인", "ForeignNet5DB", 65, true), ("기관", "InstNet5DB", 65, true), ("합산", "TotalNet5DB", 65, true));
            sp.Controls.Add(WCard("SECTOR TREND", _gSector, 0, 0, 0, 12), 0, 0);

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

            var hdr = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = Color.White };
            hdr.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pen = new Pen(GRID_LN)) e.Graphics.DrawLine(pen, 16, 43, hdr.Width - 16, 43);
                TextRenderer.DrawText(e.Graphics, title, new Font("Segoe UI", 9f, FontStyle.Bold),
                    new Rectangle(16, 0, hdr.Width - 32, 44), TXT_MAIN, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
            };

            inner.Dock = DockStyle.Fill;
            card.Controls.Add(inner);
            card.Controls.Add(hdr);

            card.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pth = RRect(new Rectangle(0, 0, card.Width - 1, card.Height - 1), 10))
                using (var pen = new Pen(CARD_BRD))
                    e.Graphics.DrawPath(pen, pth);
            };

            p.Controls.Add(card);
            return p;
        }

        static GraphicsPath RRect(Rectangle bounds, int radius)
        {
            int diameter = radius * 2; Size size = new Size(diameter, diameter); Rectangle arc = new Rectangle(bounds.Location, size); GraphicsPath path = new GraphicsPath();
            if (radius == 0) { path.AddRectangle(bounds); return path; }
            path.AddArc(arc, 180, 90); arc.X = bounds.Right - diameter; path.AddArc(arc, 270, 90);
            arc.Y = bounds.Bottom - diameter; path.AddArc(arc, 0, 90); arc.X = bounds.Left; path.AddArc(arc, 90, 90);
            path.CloseFigure(); return path;
        }

        // ═══════════════ 이벤트 ═══════════════════════════════

        async void BtnLogin_Click(object s, EventArgs e)
        {
            if (_ax == null) return;
            _btnLogin.Enabled = false; _lblLogin.Text = "● 연결 중..."; _lblLogin.ForeColor = AMBER;
            try
            {
                using (var cl = new KiwoomClient(_ax))
                {
                    if (await cl.LoginAsync())
                    {
                        var nm = _ax.GetLoginInfo("USER_NAME");
                        var sv = _ax.GetLoginInfo("GetServerGubun") == "1" ? "모의" : "실";
                        _lblLogin.Text = $"● {nm} ({sv})"; _lblLogin.ForeColor = GREEN;
                        _btnLogin.Bg = GREEN; _btnLogin.Text = "연결됨"; _btnLogin.Invalidate();
                        await LoadConds(cl); UpdRun();
                        _lblProg.Text = "키움 API 로그인 성공. 종목을 선택하거나 검색을 실행하세요.";
                    }
                }
            }
            catch (Exception ex)
            {
                _lblLogin.Text = "● 연결 실패"; _lblLogin.ForeColor = CORAL;
                _btnLogin.Enabled = true; _btnLogin.Bg = TEAL; _btnLogin.Invalidate();
                MessageBox.Show("로그인 실패: " + ex.Message);
            }
        }

        async Task LoadConds(KiwoomClient cl) { try { var ls = await cl.GetConditionListAsync(); _cbCond.Items.Clear(); _cbCond.Items.Add(new CI("", "— 조건 선택 —")); foreach (var (i, n) in ls) _cbCond.Items.Add(new CI(i, n)); if (_cbCond.Items.Count > 0) _cbCond.SelectedIndex = 0; } catch { } }

        void BtnCsv_Click(object s, EventArgs e) { using (var d = new OpenFileDialog { Filter = "CSV|*.csv|All|*.*" }) { if (d.ShowDialog() != DialogResult.OK) return; try { LoadCds(CsvCodeExtractor.Extract(d.FileName)); _lblCsv.Text = Path.GetFileName(d.FileName); _lblCsv.ForeColor = TEAL_D; } catch (Exception ex) { MessageBox.Show("CSV 오류: " + ex.Message); } } }

        async void CbCond_Changed(object s, EventArgs e) { if (_cbCond.SelectedItem is CI ci && ci.Idx != "") { if (_ax == null) return; try { using (var c = new KiwoomClient(_ax)) LoadCds(await c.GetConditionCodesAsync(ci.Idx, ci.Nm)); } catch (Exception ex) { MessageBox.Show("조건검색 실패: " + ex.Message); } } }

        void LoadCds(List<string> c) { _codes = c; _gStock.Rows.Clear(); foreach (var x in c) _gStock.Rows.Add("—", x, ""); UpdRun(); _kpiTotal.Val = c.Count.ToString(); _kpiTotal.Invalidate(); }
        void UpdRun() { _btnRun.Enabled = _ax?.GetConnectState() == 1 && _codes.Count > 0; }

        async void BtnRun_Click(object s, EventArgs e)
        {
            if (_running || _codes.Count == 0) return;
            SetRun(true); _res.Clear(); _gResult.Rows.Clear();
            _cts = new CancellationTokenSource();
            var eng = new AnalysisEngine(_ax);
            eng.Progress += (cur, tot, nm) => InvUI(() => { _bar.Value = (int)((double)cur / tot * 100); _bar.Invalidate(); _lblProg.Text = $"분석 진행 중... {nm} ({cur}/{tot})"; });
            eng.Log += m => System.Diagnostics.Debug.WriteLine(m);
            try
            {
                var (r, sk, sd) = await eng.RunAsync(_codes, ScoreConfig.Instance, _cts.Token);
                _res = r; _sK = sk; _sD = sd;
                FillResult(); FillSector(); FillStock(); UpdateKpi();
            }
            catch (OperationCanceledException) { _lblProg.Text = "분석이 중지되었습니다."; }
            catch (Exception ex) { MessageBox.Show("분석 오류: " + ex.Message); _lblProg.Text = "분석 중 오류 발생."; }
            finally { SetRun(false); }
        }

        void GSel(object s, EventArgs e) { if (_gStock.SelectedRows.Count == 0) return; var cd = _gStock.SelectedRows[0].Cells["Code"].Value?.ToString(); ShowDetail(_res.FirstOrDefault(r => r.Code == cd)); for (int i = 0; i < _gResult.Rows.Count; i++) if (_gResult.Rows[i].Cells["Code2"]?.Value?.ToString() == cd) { _gResult.Rows[i].Selected = true; break; } }
        void GRSel(object s, EventArgs e) { if (_gResult.SelectedRows.Count == 0) return; ShowDetail(_res.FirstOrDefault(r => r.Code == _gResult.SelectedRows[0].Cells["Code2"]?.Value?.ToString())); }

        void GRFmt(object s, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _res.Count) return;
            var r = _res[e.RowIndex]; var col = _gResult.Columns[e.ColumnIndex].Name;
            if (col == "SupplyTrendStr") e.CellStyle.ForeColor = r.SupplyTrend == SupplyTrend.상승 || r.SupplyTrend == SupplyTrend.상승반전 ? GREEN : r.SupplyTrend == SupplyTrend.하락 ? CORAL : r.SupplyTrend == SupplyTrend.하락반전 ? AMBER : TXT_MUTE;
            if (col == "TotalScore") e.CellStyle.ForeColor = r.TotalScore >= 80 ? GREEN : r.TotalScore >= 50 ? TEAL_D : r.TotalScore >= 30 ? AMBER : CORAL;
            if ((col == "ForeignNet5D" || col == "InstNet5D") && e.Value != null) { var v = e.Value.ToString(); if (v.StartsWith("+")) e.CellStyle.ForeColor = GREEN; else if (v.StartsWith("-")) e.CellStyle.ForeColor = CORAL; }
        }

        // ── 데이터 갱신 ──

        void FillResult() { _gResult.Rows.Clear(); for (int i = 0; i < _res.Count; i++) { var r = _res[i]; _gResult.Rows.Add(i + 1, r.Name, r.TotalScore.ToString("F1"), r.StockSupplyScore.ToString("F1"), FN(r.ForeignNet5D), FN(r.InstNet5D), r.SupplyTrend.ToString(), r.SectorName, r.SectorSupplyScore.ToString("F1"), r.Code); } }
        void FillSector() { _gSector.Rows.Clear(); foreach (var x in _sK.Concat(_sD).OrderByDescending(x => x.TotalNet5D)) _gSector.Rows.Add(x.SectorName, x.Market, FA(x.ForeignNet5D), FA(x.InstNet5D), FA(x.TotalNet5D)); }
        void FillStock() { _gStock.Rows.Clear(); foreach (var c in _codes) { var r = _res.FirstOrDefault(x => x.Code == c); _gStock.Rows.Add(r?.Name ?? c, c, r?.Market ?? ""); } }

        void UpdateKpi()
        {
            _kpiTotal.Val = _res.Count.ToString(); _kpiTotal.Invalidate();
            if (_res.Count > 0) { _kpiValue.Val = _res.Average(r => r.TotalScore).ToString("F1"); _kpiValue.Invalidate(); }
            _kpiSupply.Val = _res.Count(r => r.StockSupplyScore >= 50).ToString(); _kpiSupply.Invalidate();
            var sectors = _res.Select(r => r.SectorName).Where(s => !string.IsNullOrEmpty(s)).Distinct().Count();
            _kpiSector.Val = sectors.ToString(); _kpiSector.Invalidate();
        }

        // ── 세부 패널 ──

        void ShowDetail(AnalysisResult r)
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

            fl.Controls.Add(DR("외국인 당일", FN(r.ForeignNetD1), NC(r.ForeignNetD1)));
            fl.Controls.Add(DR("외국인 5일", FN(r.ForeignNet5D), NC(r.ForeignNet5D)));
            fl.Controls.Add(DR("외국인 20일", FN(r.ForeignNet20D), NC(r.ForeignNet20D)));
            fl.Controls.Add(DR("기관 당일", FN(r.InstNetD1), NC(r.InstNetD1)));
            fl.Controls.Add(DR("기관 5일", FN(r.InstNet5D), NC(r.InstNet5D)));
            fl.Controls.Add(DR("기관 20일", FN(r.InstNet20D), NC(r.InstNet20D)));
            fl.Controls.Add(DH());

            fl.Controls.Add(DR("회전율20일", r.Turnover20D.ToString("P2")));
            fl.Controls.Add(DR("회전율60일", r.Turnover60D.ToString("P2")));
            fl.Controls.Add(DR("회전율추세", r.TurnoverRate.ToString("+0.0;-0.0") + "%", r.TurnoverRate > 0 ? GREEN : r.TurnoverRate < 0 ? CORAL : TXT_SEC));
            fl.Controls.Add(DR("수급추세", r.SupplyTrend.ToString(), r.SupplyTrend == SupplyTrend.상승 || r.SupplyTrend == SupplyTrend.상승반전 ? GREEN : r.SupplyTrend == SupplyTrend.하락 ? CORAL : AMBER, true));
            _pDetail.Controls.Add(fl);
        }

        // ═══════════════ 팩토리 ═══════════════════════════════

        static DataGridView MkGrid(params (string h, string n, int w, bool r)[] cols)
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = CARD_BG,
                BorderStyle = BorderStyle.None,
                GridColor = GRID_LN,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
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
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = CARD_BG, ForeColor = TXT_MAIN, SelectionBackColor = GRID_SEL, SelectionForeColor = TXT_MAIN, Font = new Font("Segoe UI", 8.5f), Padding = new Padding(8, 0, 8, 0) },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = GRID_HDR, ForeColor = TXT_SEC, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), SelectionBackColor = GRID_HDR, Padding = new Padding(8, 0, 8, 0) },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = GRID_ALT, ForeColor = TXT_MAIN },
            };
            foreach (var (h, n, w, r2) in cols)
                g.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = h, Name = n, MinimumWidth = w, DefaultCellStyle = new DataGridViewCellStyle { Alignment = r2 ? DataGridViewContentAlignment.MiddleRight : DataGridViewContentAlignment.MiddleLeft } });
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

        static string FN(long v) => v.ToString("+#,0;-#,0;0");
        static string FA(double v) => (v / 1e8).ToString("+#,0.0억;-#,0.0억;0억");
        static Color NC(long v) => v >= 0 ? GREEN : CORAL;
        static Color SC(double s) => s >= 80 ? GREEN : s >= 50 ? TEAL_D : s >= 30 ? AMBER : CORAL;

        void SetRun(bool v) { _running = v; _btnRun.Enabled = !v && _codes.Count > 0; _btnStop.Enabled = v; _btnCsv.Enabled = !v; if (!v) { _bar.Value = 0; _bar.Invalidate(); _lblProg.Text = "분석 완료!"; } }
        void InvUI(Action a) { if (InvokeRequired) Invoke(a); else a(); }
        protected override void OnFormClosing(FormClosingEventArgs e) { _cts?.Cancel(); base.OnFormClosing(e); }

        sealed class CI { public string Idx, Nm; public CI(string i, string n) { Idx = i; Nm = n; } public override string ToString() => Nm; }
    }
}