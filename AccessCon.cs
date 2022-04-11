using System.Data;
using System.Data.OleDb;

namespace ExcelAddIn1
{
    internal class AccessCon : WorkbookOperator
    {
        public async void frmAccessCon_Load(string path, bool mdb)
        {
            //SQL作成
            using (DataTable resultDt = new DataTable())
            //Access接続準備
            using (OleDbCommand command = new OleDbCommand())
            using (OleDbDataAdapter da = new OleDbDataAdapter())
            using (OleDbConnection cnAccess = new OleDbConnection())
            {
                var builder = new System.Data.OleDb.OleDbConnectionStringBuilder();
                builder["Provider"] = mdb ? "Microsoft.Jet.OLEDB.4.0" : "Microsoft.ACE.OLEDB.12.0";
                builder["Data Source"] = path;
                //builder["Jet OLEDB:Database Password"] = "acbdefg";

                await this.DoTask(() =>
                {
                    var sql = "SELECT * FROM 社員";
                    cnAccess.ConnectionString = builder.ConnectionString;

                    //Access接続開始
                    cnAccess.Open();
                    command.Connection = cnAccess;
                    command.CommandText = sql.ToString();
                    da.SelectCommand = command;
                    //SQL実行 結果をデータテーブルに格納
                    da.Fill(resultDt);

                    object[,] vs = new object[resultDt.Rows.Count + 1, resultDt.Columns.Count];

                    int columncount = 0;
                    foreach (DataColumn column in resultDt.Columns)
                    {
                        vs[0, columncount++] = column.ColumnName;
                    }
                    //結果出力
                    for (int rowindex = 0; rowindex < resultDt.Rows.Count; rowindex++)
                    {
                        for (int colindex = 0; colindex < resultDt.Columns.Count; colindex++)
                        {
                            vs[rowindex + 1, colindex] = resultDt.Rows[rowindex][colindex];
                        }
                    }
                    SetActiveSheet_DataFormula(vs);
                });
            }
        }
    }
}