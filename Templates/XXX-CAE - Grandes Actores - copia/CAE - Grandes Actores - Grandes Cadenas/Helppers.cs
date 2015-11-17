using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace CAE___Grandes_Actores___Grandes_Cadenas
{
    public static class Helppers
    {
        public static Excel.QueryTable importData(String sConnectionString, String sSqlQuery, Microsoft.Office.Tools.Excel.WorksheetBase oWSB)
        {
            Boolean bDone = false;
            Excel.QueryTable oTable = null;
            try
            {
                 
                Excel.Range oRange = oWSB.Range["A1"];
                oTable = oWSB.QueryTables.Add(sConnectionString, oRange);
                oTable.CommandType = Excel.XlCmdType.xlCmdSql;
                oTable.CommandText = sSqlQuery;
                oTable.Refresh();
                bDone = true;

                for (int i = 1; i <= Globals.ThisWorkbook.Connections.Count; i++)
                    Globals.ThisWorkbook.Connections[i].Delete();
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
                bDone = false;
                oTable = null;
            }

            return oTable;
        }

    }
}
