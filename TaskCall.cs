using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    /// <summary>
    /// 
    /// </summary>
    public static class TaskCall
    {
        static object lockg = new object();
        public static async Task DoTask(this IDisposable error, System.Action action, bool Interactive = false, bool ScreenUpdating = false, bool EnableEvents = false)
        {
            var application = Globals.ThisAddIn.Application;
            try
            {
                if (!ScreenUpdating) application.ScreenUpdating = ScreenUpdating;
                if (!Interactive)
                {
                    application.Cursor = XlMousePointer.xlWait;
                    application.Calculation = XlCalculation.xlCalculationManual;
                    application.Interactive = Interactive;
                }
                if (!EnableEvents) application.EnableEvents = EnableEvents;
                await Task.Run(() => { lock (lockg) action(); });
            }
            catch (Exception ex)
            {
                error.Dispose();
                Trace.WriteLine(DateTime.Now.ToString("yyyy年MM月dd日 HH時mm分ss秒"), "TIME");
                Trace.WriteLine(ex.Message, "ERROR");
                Trace.WriteLine(ex.StackTrace, "TRACE");
                MessageBox.Show("エラーが発生しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (!ScreenUpdating) application.ScreenUpdating = true;
                if (!Interactive)
                {
                    application.Cursor = XlMousePointer.xlDefault;
                    application.Calculation = XlCalculation.xlCalculationAutomatic;
                    application.Interactive = true;
                }
                if (!EnableEvents) application.EnableEvents = true;
            }
        }
    }
}