using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Top50FarmaciasParafarmacias
{
    public partial class Report
    {
        private void Hoja2_Startup(object sender, System.EventArgs e)
        {
            ConfigurationHelpper oCfg = Globals.ThisWorkbook.oCfg;
            IMSClasses.Jobs.Job oJob = Globals.ThisWorkbook.oJob;
            IMSClasses.DBHelper.db oDB = Globals.ThisWorkbook.oDb;

            String sOutputPath = System.IO.Path.Combine(oCfg.ProccesPath, oJob.JOBCODE);
            sOutputPath = System.IO.Path.Combine(sOutputPath, oCfg.OutputPath);
            oJob.OutputParameters.SetupPath(sOutputPath);

            try
            {
                if(Globals.ThisWorkbook.StatusCorrect)
                {
                    //set up data on template
                    if (this.paint_data())
                    {
                        if (this.format_data() && this.SetUpTittle())
                        {//hide data
                            String sFileName = oJob.OutputParameters.DestinationFile.FileName;
                            System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("es-ES");
                            String MonthName = ci.DateTimeFormat.GetMonthName(DateTime.Now.Month) + " " + DateTime.Now.Year.ToString();
                            //sFileName = sFileName.Replace("%date%", DateTime.Now.ToString("yyyyMMddhhmmss"));
                            MonthName = Capitalize(MonthName);
                            sFileName = sFileName.Replace("%date%", MonthName);

                            Globals.ImportData.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                            if (System.IO.File.Exists(sFileName)) System.IO.File.Delete(sFileName);
                            Globals.ThisWorkbook.SaveAs(sFileName);


                            oJob.ReportStatus.Message = "File Created Correctly.";
                            oJob.ReportStatus.Status = "DONE";
                            oDB.updateJob(oJob.Serialize(), oJob.JOBID);
                        }
                        else
                        {
                            Globals.ThisWorkbook.StatusCorrect = false;
                            Globals.ThisWorkbook.StatusMessage = "Error formating data.";

                            oJob.ReportStatus.Message = "Error formating data.";
                            oJob.ReportStatus.Status = "ERRO";
                            oDB.updateJob(oJob.Serialize(), oJob.JOBID);
                        }
                    }
                    else
                    {
                        Globals.ThisWorkbook.StatusCorrect = false;
                        Globals.ThisWorkbook.StatusMessage = "Error Writting Report Data";

                        oJob.ReportStatus.Message = Globals.ThisWorkbook.StatusMessage;
                        oJob.ReportStatus.Status = "ERRO";
                        oDB.updateJob(oJob.Serialize(), oJob.JOBID);
                    }

                }
            }
            catch (Exception eReport)
            {
                Globals.ThisWorkbook.StatusCorrect = false;
                Globals.ThisWorkbook.StatusMessage = "Error(Not Controled)--> Sheet Report --> Exception Message: " + eReport.Message.ToString();

                oJob.ReportStatus.Message = Globals.ThisWorkbook.StatusMessage;
                oJob.ReportStatus.Status = "ERRO";
                oDB.updateJob(oJob.Serialize(), oJob.JOBID);
            }
            
        }

        private void Hoja2_Shutdown(object sender, System.EventArgs e)
        {
        }

        public String Capitalize(String sValue)
        {
            string sResult = "";

            if (sValue.Length > 1)
                sResult = char.ToUpper(sValue[0]) + sValue.Substring(1);

            return sResult;
        }

        public bool SetUpTittle()
        {
            //////////test
            String sExpresionMat = @"MAT\/[0-9][0-9]*\/[0-9]*";
            String sExpresionYtd = @"YTD\/[0-9][0-9]*\/[0-9]*";
            System.Text.RegularExpressions.Regex oRegexMat = new System.Text.RegularExpressions.Regex(sExpresionMat);
            System.Text.RegularExpressions.Regex oRegexYtd = new System.Text.RegularExpressions.Regex(sExpresionYtd);
            String ColText = "";
            // MAT
            //B1 = J4
            //E1 = M4
            //N1 = W4
            //Q1 = Z4
            List<KeyValuePair<string, string>> oList = new List<KeyValuePair<string, string>>();
            oList.Add(new KeyValuePair<string, string>("B1", "J4"));
            oList.Add(new KeyValuePair<string, string>("E1", "M4"));
            oList.Add(new KeyValuePair<string, string>("N1", "W4"));
            oList.Add(new KeyValuePair<string, string>("Q1", "Z4"));

            foreach (KeyValuePair<string, string> oPair in oList)
            {
                ColText = Globals.ImportData.Range[oPair.Key].Value;
                System.Text.RegularExpressions.Match oMath = oRegexMat.Match(ColText);
                if (oMath.Success)
                    this.Range[oPair.Value].Value = oMath.Value;
            }

            // YTD
            //H1 = D4
            //K1 = G4
            //T1 = Q4
            //W1 = T4
            oList.Clear();
            oList.Add(new KeyValuePair<string, string>("H1", "D4"));
            oList.Add(new KeyValuePair<string, string>("K1", "G4"));
            oList.Add(new KeyValuePair<string, string>("T1", "Q4"));
            oList.Add(new KeyValuePair<string, string>("W1", "T4"));

            foreach (KeyValuePair<string, string> oPair in oList)
            {
                ColText = Globals.ImportData.Range[oPair.Key].Value;
                System.Text.RegularExpressions.Match oMath = oRegexYtd.Match(ColText);
                if (oMath.Success)
                    this.Range[oPair.Value].Value = oMath.Value;
            }

            return true;
        }

        public bool format_data()
        {
            Top50FarmaciasParafarmacias.ImportData oData = Globals.ImportData;
            ArrayList oRangeList = new ArrayList();

            //oRangeListBorder.Add(oRangeColumn);
            oRangeList.Add(this.Range["C5:O5"]);
            oRangeList.Add(this.Range["Q5:AB5"]);
            oRangeList.Add(this.Range[this.Range["C5"], this.Range["C5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["D5:F5"], this.Range["D5:F5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["G5:I5"], this.Range["G5:I5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["J5:L5"], this.Range["J5:L5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["M5:O5"], this.Range["M5:O5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["Q5:S5"], this.Range["Q5:S5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["T5:V5"], this.Range["T5:V5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["W5:Y5"], this.Range["W5:Y5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["Z5:AB5"], this.Range["Z5:AB5"].End[Excel.XlDirection.xlDown]]);


            //oRangeColumn.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            foreach (Excel.Range oRange in oRangeList)
            {
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = 0;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = 0;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = 0;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = 0;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
            }


            oRangeList.Clear();
            //COLORES
            oRangeList.Add(this.Range["C5:O5"]);
            oRangeList.Add(this.Range["Q5:AB5"]);
            foreach (Excel.Range oRangeColor in oRangeList)
            {
                oRangeColor.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                oRangeColor.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                oRangeColor.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorLight2;
                oRangeColor.Interior.TintAndShade = 0.599993896298105;
                oRangeColor.Interior.PatternTintAndShade = 0;

                oRangeColor.Font.Bold = true;
            }

            //negritas para los nombres de laboratorio
            this.Range[this.Range["C5"], this.Range["C5"].End[Excel.XlDirection.xlDown]].Font.Bold = true;

            oRangeList.Clear();

            Excel.Range oRangeVichy = this.Cells.Find(What: "VICHY");
            Excel.Range oRangeRoche = this.Cells.Find(What: "LA ROCHE POSAY");
            oRangeVichy = this.Cells[oRangeVichy.Row, oRangeVichy.Column];
            oRangeList.Add(this.Range[oRangeVichy, this.Range["O" + oRangeVichy.Row.ToString()]]);
            oRangeList.Add(this.Range[this.Range["Q" + oRangeVichy.Row.ToString()], this.Range["AB" + oRangeVichy.Row.ToString()]]);

            oRangeRoche = this.Cells[oRangeRoche.Row, oRangeRoche.Column];
            oRangeList.Add(this.Range[oRangeRoche, this.Range["O" + oRangeRoche.Row.ToString()]]);
            oRangeList.Add(this.Range[this.Range["Q" + oRangeRoche.Row.ToString()], this.Range["AB" + oRangeRoche.Row.ToString()]]);


            foreach (Excel.Range oRangeFind in oRangeList)
            {
                //oRangeFind.Value = "test";
                oRangeFind.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                oRangeFind.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3;
                oRangeFind.Interior.TintAndShade = 0.399975585192419;
                oRangeFind.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            }




            oRangeList.Clear();


            oRangeList.Add(this.Range[this.Range["C6:O6"], this.Range["C" + (oRangeVichy.Row - 1).ToString() + ":O" + (oRangeVichy.Row - 1).ToString()]]);
            oRangeList.Add(this.Range[this.Range["Q6:AB6"], this.Range["Q" + (oRangeVichy.Row - 1).ToString() + ":AB" + (oRangeVichy.Row - 1).ToString()]]);

            Excel.Range oRangeAux = this.Range["C" + (oRangeRoche.Row + 1).ToString() + " :O" + (oRangeRoche.Row + 1).ToString()];
            oRangeList.Add(this.Range[oRangeAux, oRangeAux.End[Excel.XlDirection.xlDown]]);
            oRangeAux = this.Range["Q" + (oRangeRoche.Row + 1).ToString() + " :AB" + (oRangeRoche.Row + 1).ToString()];
            oRangeList.Add(this.Range[oRangeAux, oRangeAux.End[Excel.XlDirection.xlDown]]);
            /*
            oRangeList.Add(this.Range[this.Range["C6:O6"], this.Range["C6:O6"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["Q6:AB6"], this.Range["Q6:AB6"].End[Excel.XlDirection.xlDown]]);
            */

            foreach (Excel.Range oAlternative in oRangeList)
            {
                Excel.FormatCondition oCondition = oAlternative.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: "=RESIDUO(FILA();2)<>0");
                oCondition.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                oCondition.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5;
                oCondition.Interior.TintAndShade = 0.79998168889431;
                oCondition.StopIfTrue = false;
            }



            //Formateo de numeros
            oRangeList.Clear();
            oRangeList.Add(this.Range[this.Range["D5"], this.Range["D5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["G5"], this.Range["G5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["J5"], this.Range["J5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["M5"], this.Range["M5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["Q5"], this.Range["Q5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["T5"], this.Range["T5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["W5"], this.Range["W5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["Z5"], this.Range["Z5"].End[Excel.XlDirection.xlDown]]);
            foreach (Excel.Range oRangeNumber in oRangeList)
            {
                oRangeNumber.NumberFormat = "#,##0";
                oRangeNumber.Font.Name = "Arial";
                oRangeNumber.Font.Size = 10;
                oRangeNumber.Font.ThemeFont = Excel.XlThemeFont.xlThemeFontNone;
                oRangeNumber.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oRangeNumber.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

            oRangeList.Clear();
            oRangeList.Add(this.Range[this.Range["K5"], this.Range["K5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["N5"], this.Range["N5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["R5"], this.Range["R5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["U5"], this.Range["U5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["X5"], this.Range["X5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["AA5"], this.Range["AA5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["E5"], this.Range["E5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["H5"], this.Range["H5"].End[Excel.XlDirection.xlDown]]);
            foreach (Excel.Range oRangeNumber in oRangeList)
            {
                oRangeNumber.NumberFormat = "0.0_ ;[Red]-0.0";
                oRangeNumber.Font.Name = "Arial";
                oRangeNumber.Font.Size = 10;
                oRangeNumber.Font.ThemeFont = Excel.XlThemeFont.xlThemeFontNone;
                //oRangeNumber.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oRangeNumber.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

            oRangeList.Clear();
            oRangeList.Add(this.Range[this.Range["F5"], this.Range["F5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["I5"], this.Range["I5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["L5"], this.Range["L5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["O5"], this.Range["O5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["S5"], this.Range["S5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["V5"], this.Range["V5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["Y5"], this.Range["Y5"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["AB5"], this.Range["AB5"].End[Excel.XlDirection.xlDown]]);
            foreach (Excel.Range oRangeNumber in oRangeList)
            {
                oRangeNumber.NumberFormat = "0.0";
                oRangeNumber.Font.Name = "Arial";
                oRangeNumber.Font.Size = 10;
                oRangeNumber.Font.ThemeFont = Excel.XlThemeFont.xlThemeFontNone;
                oRangeNumber.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oRangeNumber.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }


            // SETUP WITH
            Excel.Range oColumns1 = this.Columns["D:O"];
            Excel.Range oColumns2 = this.Columns["Q:AB"];
            oColumns1.ColumnWidth = 7.38;
            oColumns2.ColumnWidth = 7.38;

            this.Columns["Q"].ColumnWidth = 10;
            this.Columns["T"].ColumnWidth = 10;
            this.Columns["W"].ColumnWidth = 10;
            this.Columns["Z"].ColumnWidth = 10;



            this.Range["D5"].Activate();


            return true;
        }
        public bool paint_data()
        {
            Top50FarmaciasParafarmacias.ImportData oData = Globals.ImportData;

            //Copy laboratories (A2:End) (C5)
            Excel.Range oRangeStart = oData.Range["A2"];
            Excel.Range oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            Excel.Range oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["C5"]);

            //setup max length
            int iMaxRow = 0;
            iMaxRow = oRangeEnd.Row;

            // Sales Units MAT/5/2014 (Thousands)
            //oRangeStart = oData.Range["B2"];
            //oRangeEnd = oData.Range["B" + iMaxRow.ToString()]; //oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["B2:B" + iMaxRow.ToString()];//oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["J5"]);

            // Sales Units MAT/5/2014 %PPG Previous Year (Absolute)
            //oRangeStart = oData.Range["C2"];
            //oRangeEnd = oData.Range["C" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["C2:C" + iMaxRow.ToString()];//oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["K5"]);

            // Sales Units MAT/5/2014 %V (Absolute)	L5	
            //oRangeStart = oData.Range["D2"];
            //oRangeEnd = oData.Range["D" + iMaxRow.ToString()];////oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["D2:D" + iMaxRow.ToString()];//oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["L5"]);

            // Sales Units MAT/5/2015 (Thousands)	M5	
            //oRangeStart = oData.Range["E2"];
            //oRangeEnd = oData.Range["E" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["E2:E" + iMaxRow.ToString()];//oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["M5"]);

            // Sales Units MAT/5/2015 %PPG Previous Year (Absolute)	N5	
            //oRangeStart = oData.Range["F2"];
            //oRangeEnd = oData.Range["F" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["F2:F" + iMaxRow.ToString()];//oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["N5"]);



            // Sales Units MAT/5/2015 %V (Absolute)	O5	
            //oRangeStart = oData.Range["G2"];
            //oRangeEnd = oData.Range["G" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["G2:G" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["O5"]);

            // Sales Units YTD/5/2014 (Thousands)	D5	
            //oRangeStart = oData.Range["H2"];
            //oRangeEnd = oData.Range["H" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["H2:H" + iMaxRow.ToString()];//oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["D5"]);

            // Sales Units YTD/5/2014 %PPG Previous Year (Absolute)	E5	
            //oRangeStart = oData.Range["I2"];
            //oRangeEnd = oData.Range["I" + iMaxRow.ToString()];//RangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["I2:I" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["E5"]);

            // Sales Units YTD/5/2014 %V (Absolute)	F5	
            //oRangeStart = oData.Range["J2"];
            //oRangeEnd = oData.Range["J" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["J2:J" + iMaxRow.ToString()]; ////oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["F5"]);

            // Sales Units YTD/5/2015 (Thousands)	G5	
            oRangeStart = oData.Range["K2"];
            oRangeEnd = oData.Range["K" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["G5"]);

            // Sales Units YTD/5/2015 %PPG Previous Year (Absolute)	H5	
            //oRangeStart = oData.Range["L2"];
            //oRangeEnd = oData.Range["L" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["L2:L" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["H5"]);

            // Sales Units YTD/5/2015 %V (Absolute)	I5	
            //oRangeStart = oData.Range["M2"];
            //oRangeEnd = oData.Range["M" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["M2:M" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["I5"]);

            // Sales Euros PUB MAT/5/2014 (Thousands)	W5	
            //oRangeStart = oData.Range["N2"];
            //oRangeEnd = oData.Range["N" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["N2:N" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["W5"]);

            // Sales Euros PUB MAT/5/2014 %PPG Previous Year (Absolute)	X5	
            //oRangeStart = oData.Range["O2"];
            //oRangeEnd = oData.Range["O" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["O2:O" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["X5"]);

            // Sales Euros PUB MAT/5/2014 %V (Absolute)	Y5	
            //oRangeStart = oData.Range["P2"];
            //oRangeEnd = oData.Range["P" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["P2:P" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["Y5"]);

            // Sales Euros PUB MAT/5/2015 (Thousands)	Z5	
            //oRangeStart = oData.Range["Q2"];
            //oRangeEnd = oData.Range["Q" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["Q2:Q" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["Z5"]);

            // Sales Euros PUB MAT/5/2015 %PPG Previous Year (Absolute)	AA5	
            //oRangeStart = oData.Range["R2"];
            //oRangeEnd = oData.Range["R" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["R2:R" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["AA5"]);

            // Sales Euros PUB MAT/5/2015 %V (Absolute)	AB5	
            //oRangeStart = oData.Range["S2"];
            //oRangeEnd = oData.Range["S" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["S2:S" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["AB5"]);

            // Sales Euros PUB YTD/5/2014 (Thousands)	Q5	
            //oRangeStart = oData.Range["T2"];
            //oRangeEnd = oData.Range["T" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["T2:T" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["Q5"]);

            // Sales Euros PUB YTD/5/2014 %PPG Previous Year (Absolute)	R5	
            //oRangeStart = oData.Range["U2"];
            //oRangeEnd = oData.Range["U" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["U2:U" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["R5"]);
            // Sales Euros PUB YTD/5/2014 %V (Absolute)	S5	
            //oRangeStart = oData.Range["V2"];
            //oRangeEnd = oData.Range["V" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["V2:V" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["S5"]);

            // Sales Euros PUB YTD/5/2015 (Thousands)	T5	
            //oRangeStart = oData.Range["W2"];
            //oRangeEnd = oData.Range["W" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["W2:W" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["T5"]);
            // Sales Euros PUB YTD/5/2015 %PPG Previous Year (Absolute)	U5	
            //oRangeStart = oData.Range["X2"];
            //oRangeEnd = oData.Range["X" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["X2:X" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["U5"]);
            // Sales Euros PUB YTD/5/2015 %V (Absolute)	V5	
            //oRangeStart = oData.Range["Y2"];
            //oRangeEnd = oData.Range["Y" + iMaxRow.ToString()];//oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range["Y2:Y" + iMaxRow.ToString()]; //oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["V5"]);


            //fill white spaces.
            List<Excel.Range> lRanges = new List<Excel.Range>();
            //Excel.Range oRangeToFill = this.Range["D5:O" + iMaxRow.ToString()];
            lRanges.Add(this.Range["D5:O" + iMaxRow.ToString()]);
            lRanges.Add(this.Range["Q5:AB" + iMaxRow.ToString()]);

            foreach (Excel.Range oRangeToFill in lRanges)
            {
                Excel.Range RangeFind = null;
                RangeFind = oRangeToFill.Find(What: "");
                while (RangeFind != null)
                {
                    this.Cells[RangeFind.Row, RangeFind.Column].Value = "--";
                    RangeFind = null;
                    RangeFind = oRangeToFill.Find(What: "");
                }
            }
            

            return true;
        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Hoja2_Startup);
            this.Shutdown += new System.EventHandler(Hoja2_Shutdown);
        }

        #endregion

    }
}
