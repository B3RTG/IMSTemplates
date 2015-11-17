using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace CAE___Granders_Actores___Referencias
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

        public static bool SetupHeaders(Microsoft.Office.Tools.Excel.WorksheetBase Data, Microsoft.Office.Tools.Excel.WorksheetBase Report)
        {

            //TITTLES
            String HeaderValue = "";
            String sExpresion = @"/[0-9][0-9][0-9][0-9]";
            System.Text.RegularExpressions.Regex oRegexMat = new System.Text.RegularExpressions.Regex(sExpresion);
            List<KeyValuePair<string, string>> oList = new List<KeyValuePair<string, string>>();
            oList.Add(new KeyValuePair<string, string>("B1", "C3"));
            oList.Add(new KeyValuePair<string, string>("AW1", "D3"));
            oList.Add(new KeyValuePair<string, string>("AZ1", "H3"));
            oList.Add(new KeyValuePair<string, string>("AN1", "AE3"));
            oList.Add(new KeyValuePair<string, string>("AQ1", "AI3"));

            foreach (KeyValuePair<string, string> oPair in oList)
            {
                HeaderValue = Data.Range[oPair.Key].Value;
                System.Text.RegularExpressions.Match oMath = oRegexMat.Match(HeaderValue);
                if (oMath.Success)
                {
                    String CurrentTittle = Report.Range[oPair.Value].Value;
                    CurrentTittle = CurrentTittle.Replace("YYYY", oMath.Value.Substring(1));
                    Report.Range[oPair.Value].Value = CurrentTittle;
                }
            }

            oList.Clear();
            oList.Add(new KeyValuePair<string, string>("AC1", "L3"));
            oList.Add(new KeyValuePair<string, string>("AE1", "O3"));
            oList.Add(new KeyValuePair<string, string>("AH1", "S3"));
            oList.Add(new KeyValuePair<string, string>("AK1", "W3"));
            sExpresion = @"/[0-9][0-9]*/[0-9][0-9][0-9][0-9]";
            System.Text.RegularExpressions.Regex oRegexMatMonth = new System.Text.RegularExpressions.Regex(sExpresion);
            foreach (KeyValuePair<string, string> oPair in oList)
            {
                HeaderValue = Data.Range[oPair.Key].Value;
                System.Text.RegularExpressions.Match oMathMonth = oRegexMatMonth.Match(HeaderValue);
                if (oMathMonth.Success)
                {
                    String sMonth = oMathMonth.Value.Substring(1, oMathMonth.Value.Substring(1).IndexOf('/'));
                    String sYear = oMathMonth.Value.Substring(oMathMonth.Value.Length - 4, 4);
                    DateTime oDate = new DateTime(int.Parse(sYear), int.Parse(sMonth), 1);
                    Report.Range[oPair.Value].Value = oDate;
                }
            }

            return true;
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

        public static void fillEmptySpacesCells(Microsoft.Office.Tools.Excel.WorksheetBase oWSB, Excel.Range rLookAt)
        {
            //oWSB.Cells.Replace(What: "", Replacement: "--");
            rLookAt.Replace(What: "", Replacement: "--");
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
