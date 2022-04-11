using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
namespace ExcelAddIn1
{
    /// <summary>
    /// とりあえず基本的な操作をまとめる
    /// </summary>
    internal class WorkbookOperator : Object, IDisposable
    {
        public const int FIRSTCELL = 1;
        ConcurrentStack<_Application> Applist = new ConcurrentStack<_Application>();
        ConcurrentStack<_Workbook> Worklist = new ConcurrentStack<_Workbook>();

        // Excelを起動する
        void NewOpenExcel()
        {
            Applist.Push(new Microsoft.Office.Interop.Excel.Application());
            Applist.First().Visible = false;
        }
        //ブック取得
        Workbook GetWorkbook(string filename, bool readOnly)
        {
            NewOpenExcel();
            Workbook wb = Applist.First().Workbooks.Open(Filename: filename, ReadOnly: readOnly);
            if (wb == null) return null;
            Worklist.Push(wb);
            return wb;
        }
        public virtual void Dispose()
        {
            foreach (var sheet in Worklist)
            {
                if (sheet == null) continue;
                sheet.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            }
            foreach (var sheet in Applist)
            {
                if (sheet == null) continue;
                sheet.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            }
            
            Applist.Clear();
            Worklist.Clear();
            GC.Collect();
        }
        //シート取得
        protected Sheets GetSheet(string filename, bool readOnly)
        {
            GetWorkbook(filename, readOnly);
            return Worklist.First().Sheets;
        }
        protected Sheets GetSameSheet()
        {
            return Worklist.First()?.Sheets;
        }
        //シート取得
        protected Worksheet GetWorksheet(string filename, bool readOnly,int index, bool continueReading = false)
        {
            if (index < 1) 
                return null;

            Sheets getsheet = continueReading ? GetSameSheet() : GetSheet(filename, readOnly);
            if(getsheet.Count < index)
                return null;

            return getsheet[index];
        }

        protected void GetWorksheet_DataFormula(string filename, bool readOnly, int index, out object[,] InputObject)
        {
            Worksheet ExcelWorksheet = GetWorksheet(filename, readOnly, index);
            if (ExcelWorksheet == null) { InputObject = null; return; }
            ExcelWorksheet.Select();
            Microsoft.Office.Interop.Excel.Range InputRange = ExcelWorksheet.UsedRange; // 使用範囲だけを選択する場合
            // 指定された範囲のセルの値をオブジェクト型の配列に読み込む
            InputObject = (System.Object[,])InputRange.Formula; // 数式を読む場合
        }
        protected IList<object[,]> GetAllWorksheet_DataFormula(string filename, bool readOnly)
        {
            List<object[,]> InputObject = new List<object[,]>();
            foreach (Worksheet ExcelWorksheet in GetSheet(filename, readOnly))
            {
                ExcelWorksheet.Select();
                Microsoft.Office.Interop.Excel.Range InputRange = ExcelWorksheet.UsedRange; // 使用範囲だけを選択する場合
                // 指定された範囲のセルの値をオブジェクト型の配列に読み込む
                InputObject.Add((System.Object[,])InputRange.Formula); // 数式を読む場合
            }
            return InputObject;
        }
        protected void SetActiveSheet_DataFormula(in object[,] InputObject)
        {
            Worksheet OutputWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            OutputWorksheet.Select();
            Microsoft.Office.Interop.Excel.Range OutputRange = OutputWorksheet.Range[OutputWorksheet.Cells[1, 1], OutputWorksheet.Cells[InputObject.GetLength(0), InputObject.GetLength(1)]];

            // 指定された範囲にオブジェクト型配列の値を書き込む
            OutputRange.Formula = InputObject;
        }
        protected void SetActiveSheet_DataFormula(in object[,] InputObject,int index)
        {
            Worksheet OutputWorksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[index];
            OutputWorksheet.Select();
            Microsoft.Office.Interop.Excel.Range OutputRange = OutputWorksheet.Range[OutputWorksheet.Cells[1, 1], OutputWorksheet.Cells[InputObject.GetLength(0), InputObject.GetLength(1)]];

            // 指定された範囲にオブジェクト型配列の値を書き込む
            OutputRange.Formula = InputObject;
        }
        protected void SetActiveSheet_DataFormula(in IList<object[,]> InputObjects)
        {
            Sheets OutputWorksheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets;
            var enm = InputObjects.GetEnumerator();
            foreach (Worksheet OutputWorksheet in OutputWorksheets)
            {
                if (!enm.MoveNext()) break;
                OutputWorksheet.Select();
                Microsoft.Office.Interop.Excel.Range OutputRange = OutputWorksheet.Range[OutputWorksheet.Cells[1, 1], OutputWorksheet.Cells[enm.Current.GetLength(0), enm.Current.GetLength(1)]];

                // 指定された範囲にオブジェクト型配列の値を書き込む
                OutputRange.Formula = enm.Current;
            }
        }
    }
}
