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

        public static void fillEmptySpaces(Microsoft.Office.Tools.Excel.WorksheetBase oWSB, int iMaxRow, List<Excel.Range> lRanges)
        {
            foreach (Excel.Range oRangeToFill in lRanges)
            {
                Excel.Range RangeFind = null;
                RangeFind = oRangeToFill.Find(What: "");
                while (RangeFind != null)
                {
                    oWSB.Cells[RangeFind.Row, RangeFind.Column].Value = "--";
                    RangeFind = null;
                    RangeFind = oRangeToFill.Find(What: "");
                }
            }
        }

        public static String Capitalize(String sValue)
        {
            string sResult = "";

            if (sValue.Length > 1)
                sResult = char.ToUpper(sValue[0]) + sValue.Substring(1);

            return sResult;
        }

    }
}
