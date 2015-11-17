using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

using Excel = Microsoft.Office.Interop.Excel;

namespace IMSClasses.excellHelpper
{
    public static class ExcelHelpper
    {

        /*private const object _UpdateLinks = 1;
        private const object _ReadOnly = true;
        private const object _Format = 5;
        private const object _Password = "";
        private const object _WriteResPassword = "";
        private const object _IgnoreReadOnlyRecommended = true;
        private const Microsoft.Office.Interop.Excel.XlPlatform _Origin = Microsoft.Office.Interop.Excel.XlPlatform.xlWindows;
        private const object _Delimiter = @"\t";
        private const object _Editable = false;
        private const object _Notify = false;
        private const object _Converter = 0;
        private const object _AddToMru = false;
        private const object _Local = 1;
        private const object _CorruptLoad = 0;*/

        public static bool executeTest(String sFile)
        {
            bool bCorrect = true;
            Excel.Application xlApp = new Excel.Application();
            
            
            xlApp.DisplayAlerts = true;
            xlApp.Visible = true;

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(
                sFile,
                0, true, 5, "", "", true,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                false, false, 0, false, 1, 0);

            xlWorkbook.Close();

            return true;
        }

        public static bool executeExcelTemplate(String ExcelFilename)
        {
            bool bCorrect = true;
            Excel.Application xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            
            try
            {
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(
                ExcelFilename,
                0, true, 5, "", "", true,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                false, false, 0, false, 1, 0);

                xlWorkbook.Close();
            }
            catch (Exception xlException)
            {
                if (xlApp != null) xlApp.Workbooks.Close();
                bCorrect = false;
            }

            
            xlApp.Workbooks.Close();
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return bCorrect;
        }

        public static DataTable getExcelData(String ExcelFileName, String TableName)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            String sFileName = ExcelFileName;//System.IO.Path.Combine(ExcelFilePath, ExcelFileName);
            Excel.Application xlApp = new Excel.Application();
            try
            {   
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(sFileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, false, 1, 0);

                Excel._Worksheet xlWorksheet = (Excel._Worksheet) xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                object[,] cellValues;
                cellValues = xlRange.Value2 as object[,];

                
                //System.Data.DataRow[] dr;

                for (int i = 1; i <= /*1277*/ rowCount; i++)
                {
                    String temp = "";

                    for (int j = 1; j <= colCount; j++)
                    {

                        if (cellValues[i, j] == null)
                        {
                            if (i == 1)
                            {
                                cellValues[i, j] = "Col" + j.ToString();
                            }
                            else
                            {
                                cellValues[i, j] = "";
                            }
                        }

                        String sValue = cellValues[i, j].ToString().Trim();
                        
                        
                        double dNumber = 0;
                        if (j > 1 && i > 1)
                        {
                            
                            if (double.TryParse(sValue, out dNumber))
                            {
                                //cellValues[i, j] = dNumber;
                                //do nothing
                            }
                            else if(sValue.Equals("---"))
                            {
                                cellValues[i, j] = "null";
                            }
                            else
                            {
                                cellValues[i, j] = "0,00";
                            }
                        }

                        /*if (cellValues[i, j].ToString().Trim() == "---" || 
                            cellValues[i, j].ToString().Trim() == "0"   ||
                            cellValues[i, j].ToString().Trim().Equals(String.Empty)
                        )  
                        {
                            cellValues[i, j] = "0,00";
                        }*/
                        if(i==1)
                            temp += cellValues[i, j].ToString().Replace("\n", " ").Replace("  ", " ") + "|";
                        else
                            temp += cellValues[i, j].ToString() + "|";
                    }
                    dt.TableName = TableName;
                    temp = temp.Substring(0, temp.Length - 1);

                    if (i == 1)
                    {
                        String[] tmp = temp.Split('|');
                        bool bFirst=true;

                        foreach (String col in tmp)
                        {
                            if (bFirst)
                            {
                                bFirst = false;
                                dt.Columns.Add(col);
                            }
                            else
                            {
                                Type t = typeof(Double);
                                dt.Columns.Add(col, t);                               
                            }
                            
                            
                        }
                    }
                    else
                    {
                        String[] row = temp.Split('|').ToArray<String>();
                        if(row.Contains<String>("null"))
                        {                            
                            for(int iRow=0;iRow < row.Length;iRow++)
                            {
                                if (row[iRow].Equals("null")) row[iRow] = null;
                            }
                        }
                        dt.Rows.Add(row);
                    }

                }
                xlWorkbook.Close(0);
                xlApp.Workbooks.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);                
                
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();

               
            }
            catch(Exception xlException)
            {
                if (xlApp != null)
                {
                    xlApp.Workbooks.Close();
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }
                throw xlException;
            }


            return dt;
        }

     
    }
}
