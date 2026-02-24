using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using StockAnalyzer.Models;

namespace StockAnalyzer.Forms
{
    public class SettingsForm : Form
    {
        static readonly Color SIDEBAR = Color.FromArgb(30, 39, 53);
        static readonly Color TEAL    = Color.FromArgb(0, 188, 180);
        static readonly Color BG      = Color.FromArgb(240, 242, 247);
        static readonly Color CARD    = Color.White;
        static readonly Color BRD     = Color.FromArgb(228, 232, 240);
        static readonly Color TXT     = Color.FromArgb(40, 48, 62);
        static readonly Color TXT2    = Color.FromArgb(120, 130, 150);
        static readonly Color CORAL   = Color.FromArgb(233, 87, 87);

        ScoreConfig _cfg;
        TableLayoutPanel _tbl;

        public SettingsForm()
        {
            _cfg = ScoreConfig.Load();
            Text = "설정"; Size = new Size(490, 660); MinimumSize = new Size(420, 520);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = BG; ForeColor = TXT; FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false;

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, BackColor = BG };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            Controls.Add(root);

            // 헤더
            var hdr = new Panel { Dock = DockStyle.Fill, BackColor = SIDEBAR };
            hdr.Paint += (s, e) => { TextRenderer.DrawText(e.Graphics, "◆  설정", new Font("Segoe UI Semibold", 10.5f), new Point(16, 13), Color.White); };
            root.Controls.Add(hdr, 0, 0);

            // 본문
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = BG, Padding = new Padding(12) };
            var cardOuter = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = CARD, Padding = new Padding(1) };
            cardOuter.Paint += (s, e) => { using (var pen = new Pen(BRD)) e.Graphics.DrawRectangle(pen, 0, 0, cardOuter.Width - 1, cardOuter.Height - 1); };

            _tbl = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(16, 8, 16, 8), BackColor = CARD, Width = 430 };
            _tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
            _tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            cardOuter.Controls.Add(_tbl);
            scroll.Controls.Add(cardOuter);
            root.Controls.Add(scroll, 0, 1);

            Sec("VALUATION — 기업가치");
            var perBox = Row("PER 할인율 만점", _cfg.PerScore);
            var pbrBox = Row("PBR 할인율 만점", _cfg.PbrScore);
            var roeBox = Row("ROE 만점", _cfg.RoeScore);

            Sec("STOCK SUPPLY — 종목수급");
            var fd1 = Row("외국인 당일", _cfg.ForeignD1Score);
            var f5 = Row("외국인 5일", _cfg.Foreign5DScore);
            var f20 = Row("외국인 20일", _cfg.Foreign20DScore);
            var id1 = Row("기관 당일", _cfg.InstD1Score);
            var i5 = Row("기관 5일", _cfg.Inst5DScore);
            var i20 = Row("기관 20일", _cfg.Inst20DScore);
            var tv = Row("거래회전율 추세", _cfg.TurnoverScore);

            Sec("SECTOR SUPPLY — 업종수급");
            var sfd1 = Row("업종 외국인 당일", _cfg.SectorForeignD1Score);
            var sf5 = Row("업종 외국인 5일", _cfg.SectorForeign5DScore);
            var sf20 = Row("업종 외국인 20일", _cfg.SectorForeign20DScore);
            var sid1 = Row("업종 기관 당일", _cfg.SectorInstD1Score);
            var si5 = Row("업종 기관 5일", _cfg.SectorInst5DScore);
            var si20 = Row("업종 기관 20일", _cfg.SectorInst20DScore);

            Sec("THRESHOLDS — 기준값");
            var thBox = Row("보합 기준 변화율 (%)", _cfg.TrendThresholdPct);
            var tfBox = Row("거래회전율 만점 (%)", _cfg.TurnoverFullPct);

            Sec("KRX OPEN API");
            var authBox = TxtRow("API 인증키", _cfg.KrxAuthKey);

            // 버튼바
            var bbar = new Panel { Dock = DockStyle.Fill, BackColor = CARD };
            bbar.Paint += (s, e) => { using (var p = new Pen(BRD)) e.Graphics.DrawLine(p, 0, 0, bbar.Width, 0); };

            var bSave   = new DkBtn("저장",   TEAL,                          Color.White, 76, 32);
            var bCancel = new DkBtn("취소",   Color.FromArgb(245, 247, 252), TXT,         76, 32);
            var bReset  = new DkBtn("기본값", Color.FromArgb(245, 247, 252), TXT2,        76, 32);
            bbar.Resize += (s, e) => { bSave.Location = new Point(bbar.Width - 92, 11); bCancel.Location = new Point(bbar.Width - 174, 11); bReset.Location = new Point(14, 11); };
            bbar.Controls.AddRange(new Control[] { bSave, bCancel, bReset });
            root.Controls.Add(bbar, 0, 2);

            bSave.Click += (s, e) =>
            {
                _cfg.PerScore = V(perBox); _cfg.PbrScore = V(pbrBox); _cfg.RoeScore = V(roeBox);
                _cfg.ForeignD1Score = V(fd1); _cfg.Foreign5DScore = V(f5); _cfg.Foreign20DScore = V(f20);
                _cfg.InstD1Score = V(id1); _cfg.Inst5DScore = V(i5); _cfg.Inst20DScore = V(i20);
                _cfg.TurnoverScore = V(tv);
                _cfg.SectorForeignD1Score = V(sfd1); _cfg.SectorForeign5DScore = V(sf5); _cfg.SectorForeign20DScore = V(sf20);
                _cfg.SectorInstD1Score = V(sid1); _cfg.SectorInst5DScore = V(si5); _cfg.SectorInst20DScore = V(si20);
                _cfg.TrendThresholdPct = V(thBox); _cfg.TurnoverFullPct = V(tfBox);
                _cfg.KrxAuthKey = authBox.Text.Trim();
                _cfg.Save(); DialogResult = DialogResult.OK; Close();
            };
            bCancel.Click += (s, e) => Close();
            bReset.Click += (s, e) => { if (MessageBox.Show("기본값으로 초기화할까요?", "확인", MessageBoxButtons.YesNo) == DialogResult.Yes) { _cfg = new ScoreConfig(); _cfg.Save(); Close(); } };
        }

        void Sec(string t)
        {
            var pnl = new Panel { Height = 28, Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 2), BackColor = Color.White };
            pnl.Paint += (s, e) =>
            {
                using (var b = new SolidBrush(TEAL)) e.Graphics.FillRectangle(b, 0, 20, 3, 8);
                TextRenderer.DrawText(e.Graphics, t, new Font("Segoe UI Semibold", 8f), new Rectangle(10, 0, 400, 28), TEAL, TextFormatFlags.VerticalCenter);
            };
            _tbl.Controls.Add(pnl); _tbl.SetColumnSpan(pnl, 2);
        }

        NumericUpDown Row(string label, double val)
        {
            _tbl.Controls.Add(new Label { Text = label, Height = 28, Dock = DockStyle.Fill, ForeColor = TXT2, Font = new Font("Segoe UI", 8.8f), TextAlign = ContentAlignment.MiddleLeft, BackColor = Color.White });
            var n = new NumericUpDown { Value = (decimal)val, Minimum = 0, Maximum = 100, DecimalPlaces = 1, Increment = 0.5m, Height = 26, Width = 90, BackColor = Color.FromArgb(248, 249, 252), ForeColor = TXT, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 8.8f) };
            _tbl.Controls.Add(n); return n;
        }

        TextBox TxtRow(string label, string val)
        {
            _tbl.Controls.Add(new Label { Text = label, Height = 28, Dock = DockStyle.Fill, ForeColor = TXT2, Font = new Font("Segoe UI", 8.8f), TextAlign = ContentAlignment.MiddleLeft, BackColor = Color.White });
            var t = new TextBox { Text = val ?? "", Height = 26, Width = 195, BackColor = Color.FromArgb(248, 249, 252), ForeColor = TXT, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Consolas", 8.5f) };
            _tbl.Controls.Add(t); return t;
        }

        static double V(NumericUpDown n) => (double)n.Value;
    }
}
