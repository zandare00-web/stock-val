using System;
using System.Windows.Forms;

namespace StockAnalyzer
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                Application.Run(new Forms.MainForm());
            }
            catch (Exception ex)
            {
                MessageBox.Show("실행 오류:\n\n" + ex.ToString(),
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
