using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private async void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //テスト用。5秒待機。
            const long fiveSecond = TimeSpan.TicksPerSecond * 5;
            long time = DateTime.Now.Ticks + fiveSecond;
            await MyTask.WaitUntil(() => time < DateTime.Now.Ticks);
            MessageBox.Show("５秒!!!", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //2シート目のデータを持ってくる
            using (Implementation workbookOperator = new Implementation())
            {
                workbookOperator.FromSheetCopy(2).Wait();
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //データベースのデータを持ってくる
            AccessCon accessCon = new AccessCon();
            string filename = Globals.ThisAddIn.Application.InputBox("データベースのパス?");
            accessCon.frmAccessCon_Load(filename, true);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //追加可能な限り、データも持ってくる
            using (Implementation workbookOperator = new Implementation())
            {
                workbookOperator.FromSheetCopy().Wait();
            }
        }
    }
}
