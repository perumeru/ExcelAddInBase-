using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn1
{
    internal class Implementation : WorkbookOperator
    {
        public async Task FromSheetCopy()
        {
            await this.DoTask(() =>
            {
                string filename = Globals.ThisAddIn.Application.InputBox("コピー元のファイル名を入力");
                var data = GetAllWorksheet_DataFormula(filename, true);
                SetActiveSheet_DataFormula(data);
            });
        }
        public async Task FromSheetCopy(int Sheetindex = 1)
        {
            await this.DoTask(() =>
            {
                object[,] InputObject;
                string filename = Globals.ThisAddIn.Application.InputBox("コピー元のファイル名を入力");
                GetWorksheet_DataFormula(filename, true, Sheetindex, out InputObject);
                SetActiveSheet_DataFormula(InputObject);
            });
        }
        public async Task FromSheetCopy(string strfind, int Sheetindex = 1)
        {
            await this.DoTask(() =>
            {
                object[,] InputObject;
                string filename = Globals.ThisAddIn.Application.InputBox("コピー元のファイル名を入力");
                GetWorksheet_DataFormula(filename, true, Sheetindex, out InputObject);
                LinkedList<int> list = new LinkedList<int>();

                for (int i = FIRSTCELL; i < InputObject.GetLength(1); i++)
                {
                    if(strfind.Equals(InputObject[FIRSTCELL, i].ToString()))
                    {
                        list.AddLast(i);
                    }
                }
                object[,] vs = new object[list.Count, InputObject.GetLength(0)];
                int count = 0;
                foreach (int index in list)
                {
                    for (int j = 0; j < InputObject.GetLength(0); j++)
                    {
                        vs[count, j] = InputObject[j + FIRSTCELL, index];
                    }
                    count++;
                }
                
                SetActiveSheet_DataFormula(vs);
            });
        }
    }
}
