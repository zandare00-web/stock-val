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
        public int Rad = 6;
        bool _h;
        public RndBtn(string t, Color bg, Color fg, int w, int h) { Text = t; Bg = bg; Fg = fg; Bdr = Color.Empty; Size = new Size(w, h); Font = new Font("Segoe UI Semibold", 8.5f); Cursor = Cursors.Hand; SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer | ControlStyles.SupportsTransparentBackColor, true); BackColor = Color.Transparent; }
        protected override void OnPaint(PaintEventArgs e) { var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias; var r = new Rectangle(0, 0, Width - 1, Height - 1); var bg = Enabled ? (_h ? Lt(Bg, 15) : Bg) : Color.FromArgb(180, 185, 195); using (var p = RR(r, Rad)) { using (var b = new SolidBrush(bg)) g.FillPath(b, p); if (Bdr != Color.Empty) using (var pen = new Pen(Bdr)) g.DrawPath(pen, p); } TextRenderer.DrawText(g, Text, Font, r, Enabled ? Fg : Color.FromArgb(140, 145, 155), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter); }
        protected override void OnMouseEnter(EventArgs e) { _h = true; Invalidate(); }
        protected override void OnMouseLeave(EventArgs e) { _h = false; Invalidate(); }
        static Color Lt(Color c, int a) => Color.FromArgb(Math.Min(255, c.R + a), Math.Min(255, c.G + a), Math.Min(255, c.B + a));
        static GraphicsPath RR(Rectangle r, int d) { var p = new GraphicsPath(); int dd = d * 2; p.AddArc(r.X, r.Y, dd, dd, 180, 90); p.AddArc(r.Right - dd, r.Y, dd, dd, 270, 90); p.AddArc(r.Right - dd, r.Bottom - dd, dd, dd, 0, 90); p.AddArc(r.X, r.Bottom - dd, dd, dd, 90, 90); p.CloseFigure(); return p; }
    }

    internal sealed class SlimBar : Control
    {
        public int Value; public Color Bar = Color.FromArgb(0, 188, 180), Track = Color.FromArgb(220, 225, 232);
        public SlimBar() { Height = 5; SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true); }
        protected override void OnPaint(PaintEventArgs e) { var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias; int r = Height / 2; using (var b = new SolidBrush(Track)) using (var p = P(0, 0, Width, Height, r)) g.FillPath(b, p); int w = (int)(Width * Math.Max(0, Math.Min(100, Value)) / 100.0); if (w > 2) using (var b = new SolidBrush(Bar)) using (var p = P(0, 0, w, Height, r)) g.FillPath(b, p); }
        static GraphicsPath P(int x, int y, int w, int h, int r) { var p = new GraphicsPath(); if (w <= 0) return p; r = Math.Min(r, Math.Min(w / 2, h / 2)); int d = r * 2; p.AddArc(x, y, d, d, 180, 90); p.AddArc(x + w - d, y, d, d, 270, 90); p.AddArc(x + w - d, y + h - d, d, d, 0, 90); p.AddArc(x, y + h - d, d, d, 90, 90); p.CloseFigure(); return p; }
    }

    // KPI 카드 (흰 카드 위에 라벨+큰 숫자)
    internal sealed class KpiCard : Panel
    {
        public string Title = "", Val = "", Sub = "";
        public Color ValColor = Color.FromArgb(40, 48, 62);
        public KpiCard() { BackColor = Color.White; Size = new Size(150, 68); SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
            // 라운드 카드
            using (var p = RR(new Rectangle(0, 0, Width - 1, Height - 1), 8))
            { using (var b = new SolidBrush(Color.White)) g.FillPath(b, p); using (var pen = new Pen(Color.FromArgb(230, 233, 240))) g.DrawPath(pen, p); }
            // 타이틀
            TextRenderer.DrawText(g, Title, new Font("Segoe UI", 7.5f), new Rectangle(12, 8, Width - 24, 16), Color.FromArgb(130, 140, 160), TextFormatFlags.Left);
            // 큰 숫자
            TextRenderer.DrawText(g, Val, new Font("Segoe UI Semibold", 14f), new Rectangle(12, 22, Width - 24, 30), ValColor, TextFormatFlags.Left);
            // 서브텍스트
            if (!string.IsNullOrEmpty(Sub))
                TextRenderer.DrawText(g, Sub, new Font("Segoe UI", 7f), new Rectangle(12, 50, Width - 24, 14), Color.FromArgb(150, 160, 175), TextFormatFlags.Left);
        }
        static GraphicsPath RR(Rectangle r, int d) { var p = new GraphicsPath(); int dd = d * 2; p.AddArc(r.X, r.Y, dd, dd, 180, 90); p.AddArc(r.Right - dd, r.Y, dd, dd, 270, 90); p.AddArc(r.Right - dd, r.Bottom - dd, dd, dd, 0, 90); p.AddArc(r.X, r.Bottom - dd, dd, dd, 90, 90); p.CloseFigure(); return p; }
    }

    // ══════════════════════════════════════════════════════════
    //  MainForm
    // ══════════════════════════════════════════════════════════
    public partial class MainForm : Form
    {
        // ── 팔레트 (참조 이미지 기반) ──
        static readonly Color SIDEBAR   = Color.FromArgb(30, 39, 53);     // 다크 네이비 사이드바
        static readonly Color SB_SEL    = Color.FromArgb(0, 188, 180);    // 틸(Teal) 액센트
        static readonly Color SB_TXT    = Color.FromArgb(170, 180, 200);
        static readonly Color SB_TXT_A  = Color.White;
        static readonly Color MAIN_BG   = Color.FromArgb(240, 242, 247);  // 밝은 회색 본문 배경
        static readonly Color CARD_BG   = Color.White;
        static readonly Color CARD_BRD  = Color.FromArgb(228, 232, 240);
        static readonly Color HDR_BG    = Color.FromArgb(30, 39, 53);     // 헤더바
        static readonly Color HDR_TXT   = Color.White;
        static readonly Color TEAL      = Color.FromArgb(0, 188, 180);    // 메인 액센트 (틸)
        static readonly Color TEAL_D    = Color.FromArgb(0, 155, 148);
        static readonly Color CORAL     = Color.FromArgb(233, 87, 87);    // 보조 (하락)
        static readonly Color GREEN     = Color.FromArgb(38, 190, 100);   // 상승
        static readonly Color AMBER     = Color.FromArgb(245, 180, 40);
        static readonly Color TXT_MAIN  = Color.FromArgb(40, 48, 62);
        static readonly Color TXT_SEC   = Color.FromArgb(120, 130, 150);
        static readonly Color TXT_MUTE  = Color.FromArgb(165, 175, 190);
        static readonly Color GRID_HDR  = Color.FromArgb(245, 247, 252);
        static readonly Color GRID_ALT  = Color.FromArgb(250, 251, 254);
        static readonly Color GRID_LN   = Color.FromArgb(235, 238, 245);
        static readonly Color GRID_SEL  = Color.FromArgb(220, 245, 243);

        // ── 상태 ──
        AxKHOpenAPI _ax;
        List<string> _codes = new List<string>();
        List<AnalysisResult> _res = new List<AnalysisResult>();
        List<SectorSupplySummary> _sK = new List<SectorSupplySummary>(), _sD = new List<SectorSupplySummary>();
        CancellationTokenSource _cts; bool _running;

        // ── 컨트롤 ──
        RndBtn _btnLogin, _btnCsv, _btnRun, _btnStop;
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
            Text = "Stock Analyzer"; Size = new Size(1500, 920); MinimumSize = new Size(1100, 700);
            BackColor = MAIN_BG; ForeColor = TXT_MAIN; StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9f);
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);

            // ═══ 좌측 사이드바 (180px) ═══
            var sidebar = new Panel { Dock = DockStyle.Left, Width = 180, BackColor = SIDEBAR };
            sidebar.Paint += PaintSidebar;
            // 사이드바 메뉴 버튼들
            var sbItems = new[] { ("📊", "분석 대시보드", true), ("📋", "종목 관리", false), ("⚙", "설정", false) };
            int sy = 70;
            foreach (var (icon, label, active) in sbItems)
            {
                var btn = new Label
                {
                    Text = $"  {icon}  {label}", Font = new Font("Segoe UI", 9f),
                    ForeColor = active ? SB_TXT_A : SB_TXT,
                    BackColor = active ? Color.FromArgb(40, 52, 70) : SIDEBAR,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Cursor = Cursors.Hand, Height = 38, Dock = DockStyle.Top,
                    Padding = new Padding(12, 0, 0, 0),
                };
                if (active)
                {
                    btn.Paint += (s, e) =>
                    {
                        using (var b = new SolidBrush(TEAL))
                            e.Graphics.FillRectangle(b, 0, 0, 3, btn.Height);
                    };
                }
                if (label == "설정")
                    btn.Click += (s, e) => { using (var f = new SettingsForm()) f.ShowDialog(this); };
                sidebar.Controls.Add(btn);
                sidebar.Controls.SetChildIndex(btn, 0); // 위에서부터
                sy += 38;
            }
            Controls.Add(sidebar);

            // ═══ 상단 헤더바 (46px) ═══
            var header = new Panel { Dock = DockStyle.Top, Height = 46, BackColor = HDR_BG };
            header.Resize += (s, e) => LayoutHeader(header);

            var title = new Label { Text = "Stock Analyzer  ·  분석 대시보드", Font = new Font("Segoe UI Semibold", 10.5f), ForeColor = HDR_TXT, BackColor = HDR_BG, TextAlign = ContentAlignment.MiddleLeft };
            title.SetBounds(14, 0, 300, 46);

            _lblLogin = new Label { ForeColor = Color.FromArgb(255, 120, 120), Font = new Font("Segoe UI", 8.5f), BackColor = HDR_BG, TextAlign = ContentAlignment.MiddleRight, Text = "● 미연결" };
            _btnLogin = new RndBtn("연결", TEAL, Color.White, 66, 28);
            _btnLogin.Click += BtnLogin_Click;

            header.Controls.AddRange(new Control[] { title, _lblLogin, _btnLogin });
            Controls.Add(header);

            // ═══ 툴바 (42px) ═══
            var tool = new Panel { Dock = DockStyle.Top, Height = 42, BackColor = Color.White };
            tool.Paint += (s, e) => { using (var p = new Pen(CARD_BRD)) e.Graphics.DrawLine(p, 0, 41, tool.Width, 41); };
            tool.Resize += (s, e) => LayoutTool(tool);

            _btnCsv = new RndBtn("📂 CSV 불러오기", Color.FromArgb(245, 247, 252), TXT_MAIN, 120, 28) { Bdr = CARD_BRD, Font = new Font("Segoe UI", 8.5f) };
            _btnCsv.Click += BtnCsv_Click;
            _lblCsv = new Label { Text = "파일 없음", ForeColor = TXT_MUTE, Font = new Font("Segoe UI", 8f), BackColor = Color.White, TextAlign = ContentAlignment.MiddleLeft };

            _cbCond = new ComboBox { Width = 175, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.White, ForeColor = TXT_MAIN, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5f) };
            _cbCond.SelectedIndexChanged += CbCond_Changed;

            _btnRun = new RndBtn("▶  분석 시작", TEAL, Color.White, 100, 28) { Enabled = false };
            _btnRun.Click += BtnRun_Click;
            _btnStop = new RndBtn("■  중지", CORAL, Color.White, 68, 28) { Enabled = false };
            _btnStop.Click += (s, e) => _cts?.Cancel();
            _bar = new SlimBar { Width = 130 };
            _lblProg = new Label { Text = "대기 중", ForeColor = TXT_MUTE, Font = new Font("Segoe UI", 8f), BackColor = Color.White, TextAlign = ContentAlignment.MiddleLeft };

            var lcond = new Label { Text = "조건검색", ForeColor = TXT_SEC, Font = new Font("Segoe UI", 8.2f), BackColor = Color.White, TextAlign = ContentAlignment.MiddleLeft, Width = 50 };
            tool.Controls.AddRange(new Control[] { _btnCsv, _lblCsv, lcond, _cbCond, _btnRun, _btnStop, _bar, _lblProg });
            Controls.Add(tool);

            // ═══ 본문 ═══
            var body = new Panel { Dock = DockStyle.Fill, BackColor = MAIN_BG, Padding = new Padding(12) };

            // KPI 카드 행
            var kpiRow = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 78, BackColor = MAIN_BG, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(0, 0, 0, 6) };
            _kpiTotal  = new KpiCard { Title = "분석 종목수", Val = "0", Sub = "종목", Width = 155 };
            _kpiValue  = new KpiCard { Title = "평균 총점", Val = "—", Sub = "최대 125점", Width = 155 };
            _kpiSupply = new KpiCard { Title = "수급 양호 종목", Val = "0", Sub = "종목 (50점 이상)", ValColor = TEAL_D, Width = 155 };
            _kpiSector = new KpiCard { Title = "업종 수", Val = "0", Sub = "개 업종 분석", Width = 155 };
            kpiRow.Controls.AddRange(new Control[] { _kpiTotal, Sp(8), _kpiValue, Sp(8), _kpiSupply, Sp(8), _kpiSector });
            body.Controls.Add(kpiRow);

            // 3열 메인 영역
            var grid3 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, BackColor = MAIN_BG, Padding = new Padding(0, 4, 0, 0) };
            grid3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));
            grid3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 44));
            grid3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            grid3.Controls.Add(WCard("WATCHLIST — 종목 리스트", BuildP1()), 0, 0);
            grid3.Controls.Add(WCard("ANALYSIS — 분석 결과", BuildP2()), 1, 0);
            grid3.Controls.Add(BuildP3(), 2, 0);
            body.Controls.Add(grid3);
            Controls.Add(body);
        }

        // ── 레이아웃 ──

        void LayoutHeader(Panel p)
        {
            _btnLogin.Location = new Point(p.Width - 80, 9);
            _lblLogin.SetBounds(p.Width - 80 - 145, 0, 140, 46);
        }

        void LayoutTool(Panel p)
        {
            int x = 12, y = 7;
            _btnCsv.Location = new Point(x, y); x += _btnCsv.Width + 8;
            _lblCsv.SetBounds(x, 0, 90, 42); x += 94;
            var lc = p.Controls[2] as Label; lc?.SetBounds(x, 0, 50, 42); x += 52;
            _cbCond.SetBounds(x, y, 175, 28); x += 183;
            _btnRun.Location = new Point(x, y); x += _btnRun.Width + 6;
            _btnStop.Location = new Point(x, y); x += _btnStop.Width + 14;
            _bar.SetBounds(x, 19, 130, 5); x += 138;
            _lblProg.SetBounds(x, 0, 180, 42);
        }

        // ── 패널 빌더 ──

        Control BuildP1()
        {
            _gStock = MkGrid(("종목명", "Name", 120, false), ("코드", "Code", 60, false), ("시장", "Market", 45, false));
            _gStock.SelectionChanged += GSel;
            return _gStock;
        }

        Control BuildP2()
        {
            _gResult = MkGrid(
                ("#", "Rank", 28, true), ("종목", "Name", 75, false), ("총점", "TotalScore", 42, true),
                ("수급", "StockSupplyScore", 44, true), ("외국인", "ForeignNet5D", 56, true),
                ("기관", "InstNet5D", 56, true), ("추세", "SupplyTrendStr", 44, false),
                ("업종", "SectorName", 55, false), ("업종수급", "SectorSupplyScore", 48, true));
            _gResult.SelectionChanged += GRSel;
            _gResult.CellFormatting += GRFmt;
            return _gResult;
        }

        Control BuildP3()
        {
            // 세로 2분할: 업종수급 + 종목세부
            var outer = new Panel { Dock = DockStyle.Fill, BackColor = MAIN_BG, Padding = new Padding(4, 0, 0, 0) };
            var sp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = MAIN_BG };
            sp.RowStyles.Add(new RowStyle(SizeType.Percent, 42));
            sp.RowStyles.Add(new RowStyle(SizeType.Percent, 58));

            _gSector = MkGrid(("업종", "SectorName", 80, false), ("시장", "Market", 40, false),
                ("외국인", "ForeignNet5DB", 65, true), ("기관", "InstNet5DB", 65, true), ("합산", "TotalNet5DB", 65, true));
            sp.Controls.Add(WCard("SECTOR — 업종 수급 현황", _gSector), 0, 0);

            _pDetail = new Panel { Dock = DockStyle.Fill, BackColor = CARD_BG, AutoScroll = true };
            ShowDetail(null);
            sp.Controls.Add(WCard("DETAIL — 종목 세부 정보", _pDetail), 0, 1);

            outer.Controls.Add(sp);
            return outer;
        }

        // 흰색 카드 래퍼 (제목 + 테두리)
        static Panel WCard(string title, Control inner)
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = MAIN_BG, Padding = new Padding(0, 0, 4, 6) };
            var card = new Panel { Dock = DockStyle.Fill, BackColor = CARD_BG, Padding = new Padding(0) };

            // 카드 헤더 (타이틀)
            var hdr = new Panel { Dock = DockStyle.Top, Height = 32, BackColor = Color.White };
            hdr.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                // 하단 보더
                using (var pen = new Pen(CARD_BRD)) e.Graphics.DrawLine(pen, 0, 31, hdr.Width, 31);
                // 좌측 틸 인디케이터
                using (var b = new SolidBrush(TEAL)) e.Graphics.FillRectangle(b, 0, 8, 3, 16);
                // 타이틀
                TextRenderer.DrawText(e.Graphics, title, new Font("Segoe UI Semibold", 8.2f),
                    new Rectangle(12, 0, hdr.Width - 12, 32), TXT_MAIN,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
            };

            inner.Dock = DockStyle.Fill;
            card.Controls.Add(inner);
            card.Controls.Add(hdr);

            // 카드 테두리 페인트
            card.Paint += (s, e) =>
            {
                using (var pen = new Pen(CARD_BRD))
                    e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
            };

            p.Controls.Add(card);
            return p;
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
                    }
                }
            }
            catch (Exception ex)
            {
                _lblLogin.Text = "● 실패"; _lblLogin.ForeColor = CORAL;
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
            eng.Progress += (cur, tot, nm) => InvUI(() => { _bar.Value = (int)((double)cur / tot * 100); _bar.Invalidate(); _lblProg.Text = $"{nm}  {cur}/{tot}"; });
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
            if (r == null) { _pDetail.Controls.Add(new Label { Text = "종목을 선택하세요", ForeColor = TXT_MUTE, Font = new Font("Segoe UI", 9f), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.White }); return; }

            var fl = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, BackColor = Color.White, Padding = new Padding(12, 8, 12, 8) };

            fl.Controls.Add(DL($"{r.Name}  {r.Code}", new Font("Segoe UI Semibold", 10f), TXT_MAIN));
            fl.Controls.Add(DL($"{r.Market}  ·  {r.SectorName}  ·  {r.CurrentPrice:N0}원", new Font("Segoe UI", 8f), TXT_SEC));
            fl.Controls.Add(DH());

            fl.Controls.Add(DR("총점", r.TotalScore.ToString("F1"), SC(r.TotalScore)));
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
            fl.Controls.Add(DR("수급추세", r.SupplyTrend.ToString(), r.SupplyTrend == SupplyTrend.상승 || r.SupplyTrend == SupplyTrend.상승반전 ? GREEN : r.SupplyTrend == SupplyTrend.하락 ? CORAL : AMBER));
            _pDetail.Controls.Add(fl);
        }

        // ═══════════════ 팩토리 ═══════════════════════════════

        static DataGridView MkGrid(params (string h, string n, int w, bool r)[] cols)
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill, BackgroundColor = CARD_BG, BorderStyle = BorderStyle.None,
                GridColor = GRID_LN, CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                RowHeadersVisible = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeight = 30, RowTemplate = { Height = 28 },
                EnableHeadersVisualStyles = false, ScrollBars = ScrollBars.Vertical,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = CARD_BG, ForeColor = TXT_MAIN, SelectionBackColor = GRID_SEL, SelectionForeColor = TXT_MAIN, Font = new Font("Segoe UI", 8.2f), Padding = new Padding(4, 0, 4, 0) },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = GRID_HDR, ForeColor = TXT_SEC, Font = new Font("Segoe UI Semibold", 7.8f), SelectionBackColor = GRID_HDR, Padding = new Padding(4, 0, 4, 0) },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = GRID_ALT, ForeColor = TXT_MAIN },
            };
            foreach (var (h, n, w, r2) in cols)
                g.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = h, Name = n, MinimumWidth = w, DefaultCellStyle = new DataGridViewCellStyle { Alignment = r2 ? DataGridViewContentAlignment.MiddleRight : DataGridViewContentAlignment.MiddleLeft } });
            if (cols.Any(c => c.n == "Name" && cols.Any(x => x.n == "TotalScore")))
                g.Columns.Add(new DataGridViewTextBoxColumn { Name = "Code2", Visible = false });
            return g;
        }

        void PaintSidebar(object s, PaintEventArgs e)
        {
            // 로고 영역
            TextRenderer.DrawText(e.Graphics, "◆ STOCK", new Font("Segoe UI Semibold", 12f), new Rectangle(16, 14, 160, 24), Color.White, TextFormatFlags.Left);
            TextRenderer.DrawText(e.Graphics, "    ANALYZER", new Font("Segoe UI", 8.5f), new Rectangle(16, 36, 160, 18), SB_TXT, TextFormatFlags.Left);
        }

        static Panel Sp(int w) => new Panel { Width = w, Height = 1, BackColor = MAIN_BG };
        static Label DL(string t, Font f, Color c) => new Label { Text = t, AutoSize = true, Font = f, ForeColor = c, BackColor = CARD_BG, Margin = new Padding(0, 0, 0, 2) };
        static Panel DH() => new Panel { Width = 300, Height = 1, BackColor = CARD_BRD, Margin = new Padding(0, 5, 0, 5) };
        static Panel DR(string lbl, string val, Color? vc = null)
        {
            var p = new Panel { Width = 300, Height = 20, BackColor = CARD_BG };
            p.Controls.Add(new Label { Text = lbl, Width = 95, ForeColor = TXT_SEC, Font = new Font("Segoe UI", 8f), TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Left, BackColor = CARD_BG });
            p.Controls.Add(new Label { Text = val, ForeColor = vc ?? TXT_MAIN, Font = new Font("Segoe UI Semibold", 8.2f), TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill, BackColor = CARD_BG });
            return p;
        }

        static string FN(long v) => v.ToString("+#,0;-#,0;0");
        static string FA(double v) => (v / 1e8).ToString("+#,0.0억;-#,0.0억;0억");
        static Color NC(long v) => v >= 0 ? GREEN : CORAL;
        static Color SC(double s) => s >= 80 ? GREEN : s >= 50 ? TEAL_D : s >= 30 ? AMBER : CORAL;

        void SetRun(bool v) { _running = v; _btnRun.Enabled = !v && _codes.Count > 0; _btnStop.Enabled = v; _btnCsv.Enabled = !v; if (!v) { _bar.Value = 0; _bar.Invalidate(); _lblProg.Text = "완료"; } }
        void InvUI(Action a) { if (InvokeRequired) Invoke(a); else a(); }
        protected override void OnFormClosing(FormClosingEventArgs e) { _cts?.Cancel(); base.OnFormClosing(e); }

        sealed class CI { public string Idx, Nm; public CI(string i, string n) { Idx = i; Nm = n; } public override string ToString() => Nm; }
    }
}
