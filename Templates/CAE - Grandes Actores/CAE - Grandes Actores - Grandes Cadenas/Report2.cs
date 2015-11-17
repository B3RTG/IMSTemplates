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

namespace CAE___Grandes_Actores___Grandes_Cadenas
{
    public partial class Report2
    {
        private void Hoja3_Startup(object sender, System.EventArgs e)
        {
            IMSClasses.ConfigurationHelpper oCfg = Globals.ThisWorkbook.oCfg;
            IMSClasses.Jobs.Job oJob = Globals.ThisWorkbook.oJob;
            IMSClasses.DBHelper.db oDB = Globals.ThisWorkbook.oDb;

            String sOutputPath = System.IO.Path.Combine(oCfg.ProccesPath, oJob.JOBCODE);
            sOutputPath = System.IO.Path.Combine(sOutputPath, oCfg.OutputPath);
            oJob.OutputParameters.SetupPath(sOutputPath);


            try
            {
                if (Globals.ThisWorkbook.StatusCorrect)
                {
                    if (this.paint_data())
                    {
                        if (this.format_data() && SetupHeaders())
                        {
                            String sFileName = oJob.OutputParameters.DestinationFile.FileName;
                            System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("es-ES");
                            String MonthName = ci.DateTimeFormat.GetMonthName(DateTime.Now.Month) + " " + DateTime.Now.Year.ToString();
                            //sFileName = sFileName.Replace("%date%", DateTime.Now.ToString("yyyyMMddhhmmss"));
                            MonthName = Capitalize(MonthName);
                            sFileName = sFileName.Replace("%date%", MonthName);

                            Globals.Data1.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                            if (System.IO.File.Exists(sFileName)) System.IO.File.Delete(sFileName);
                            Globals.Report1.Range["A1"].Activate();
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
                    { //error pintando datos.
                        Globals.ThisWorkbook.StatusCorrect = false;
                        Globals.ThisWorkbook.StatusMessage = "Error Writting Report Data";

                        oJob.ReportStatus.Message = "Error Writting Report Data";
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
        public String Capitalize(String sValue)
        {
            string sResult = "";

            if (sValue.Length > 1)
                sResult = char.ToUpper(sValue[0]) + sValue.Substring(1);

            return sResult;
        }
        public bool format_data()
        {
            //cuadricula inicial
            Excel.Range oRangeStart = this.Range["B4:AL4"];
            Excel.Range oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            Excel.Range oRangeColumn = this.Range[oRangeStart, oRangeEnd];

            //bordes iniciales de todo
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = 0;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = 0;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = 0;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = 0;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlInsideHorizontal].ColorIndex = 0;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlInsideVertical].ColorIndex = 0;
            oRangeColumn.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;



            oRangeStart = this.Range["A4"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = this.Range[oRangeStart, oRangeEnd];

            foreach (Excel.Range oCell in this.Range[oRangeStart, oRangeEnd].Cells)
            {
                if (oCell.Value == 3 || oCell.Value == 1 || oCell.Value == 5)
                {
                    oRangeStart = this.Range["B" + oCell.Row.ToString()];
                    oRangeEnd = oRangeStart.End[Excel.XlDirection.xlToRight];
                    oRangeColumn = this.Range[oRangeStart, oRangeEnd];
                }

                if (oCell.Value == 3 || oCell.Value == 1)
                {
                    oRangeColumn.Font.Bold = true;
                    //setup borders
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = 0;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = 0;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = 0;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = 0;
                    oRangeColumn.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;

                    //setup colors
                    oRangeColumn.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    oRangeColumn.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    oRangeColumn.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorLight2;
                    oRangeColumn.Interior.TintAndShade = 0.599993896298105;
                    oRangeColumn.Interior.PatternTintAndShade = 0;
                }
                if (oCell.Value == 5)
                {
                    oRangeColumn.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    oRangeColumn.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    oRangeColumn.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3;
                    oRangeColumn.Interior.TintAndShade = 0.599993896298105;
                    oRangeColumn.Interior.PatternTintAndShade = 0;
                }
            }

            //formato masivo de numeros
            //COLUMNAS DE MILES
            ArrayList oRangeList = new ArrayList();
            oRangeList.Add(this.Range[this.Range["D4"], this.Range["D4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["H4"], this.Range["H4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["L4"], this.Range["L4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["O4"], this.Range["O4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["S4"], this.Range["S4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["W4"], this.Range["W4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["AA4"], this.Range["AA4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["AE4"], this.Range["AE4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["AI4"], this.Range["AI4"].End[Excel.XlDirection.xlDown]]);

            foreach (Excel.Range oRangeNumber in oRangeList)
            {
                oRangeNumber.NumberFormat = "#,##0";
                oRangeNumber.Font.Name = "Arial";
                oRangeNumber.Font.Size = 10;
                oRangeNumber.Font.ThemeFont = Excel.XlThemeFont.xlThemeFontNone;
                oRangeNumber.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oRangeNumber.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

            //COLUMNAS DE NEGATIVOS ROJOS
            oRangeList.Clear();
            oRangeList.Add(this.Range[this.Range["E4:G4"], this.Range["E4:G4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["I4:K4"], this.Range["I4:K4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["M4:N4"], this.Range["M4:N4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["P4:R4"], this.Range["P4:R4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["T4:V4"], this.Range["T4:V4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["X4:Z4"], this.Range["X4:Z4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["AB4:AD4"], this.Range["AB4:AD4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["AF4:AH4"], this.Range["AF4:AH4"].End[Excel.XlDirection.xlDown]]);
            oRangeList.Add(this.Range[this.Range["AJ4:AL4"], this.Range["AJ4:AL4"].End[Excel.XlDirection.xlDown]]);


            foreach (Excel.Range oRangeNumber in oRangeList)
            {
                oRangeNumber.NumberFormat = "0.0_ ;[Red]-0.0";
                oRangeNumber.Font.Name = "Arial";
                oRangeNumber.Font.Size = 10;
                oRangeNumber.Font.ThemeFont = Excel.XlThemeFont.xlThemeFontNone;
                //oRangeNumber.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oRangeNumber.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

            // grupos
            ArrayList oRangeGroups = new ArrayList();
            Excel.Range FirstGroupRange = null, LastGroupRange = null, PreviousParentRange = null, PreviousCell = null;

            oRangeStart = this.Range["A4"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = this.Range[oRangeStart, oRangeEnd];


            //bucle 
            PreviousParentRange = null;
            Excel.Range AllCellsRange = this.Range[oRangeStart, oRangeEnd];

            foreach (Excel.Range oCell in AllCellsRange.Cells)
            {
                if ((oCell.Value == 5 || oCell.Value == 3) && PreviousParentRange != null)
                { //cambio de grupo
                    Excel.Range toAdd = this.Range[this.Cells[PreviousParentRange.Row + 1, PreviousParentRange.Column], this.Cells[oCell.Row - 1, oCell.Column]];
                    oRangeGroups.Add(toAdd);
                    if (oCell.Value == 5)
                        PreviousParentRange = oCell;
                    else
                        PreviousParentRange = null;
                }
                else if (oCell.Value == 5 && PreviousParentRange == null)
                {
                    PreviousParentRange = oCell;
                }

                PreviousCell = oCell;
            }


            //bucle para raizes
            foreach (Excel.Range oCell in AllCellsRange.Cells)
            {
                if ((oCell.Value == 3 || oCell.Value == 1) && PreviousParentRange != null)
                { //cambio de grupo
                    Excel.Range toAdd = this.Range[this.Cells[PreviousParentRange.Row + 1, PreviousParentRange.Column], this.Cells[oCell.Row - 1, oCell.Column]];
                    oRangeGroups.Add(toAdd);
                    PreviousParentRange = oCell;
                }
                else if (oCell.Value == 3 || oCell.Value == 1)
                {
                    PreviousParentRange = oCell;
                }
                else if (AllCellsRange.Cells.Count == oCell.Row - 3 && PreviousParentRange != null)
                {
                    Excel.Range toAdd = this.Range[this.Cells[PreviousParentRange.Row + 1, PreviousParentRange.Column], this.Cells[oCell.Row, oCell.Column]];
                    oRangeGroups.Add(toAdd);
                }

                PreviousCell = oCell;
            }

            foreach (Excel.Range oRangeToGroup in oRangeGroups)
            {
                oRangeToGroup.Rows.Group();
            }

            oRangeStart = this.Range["A4"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = this.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Value = "";

            this.Columns["C:AL"].ColumnWidth = 9;
            this.Columns["B"].ColumnWidth = 45;
            this.Columns["C"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            Excel.Range RangeFind = null, FirstFind = null;
            bool bFirst = true;
            RangeFind = this.Cells.Find(What: "Total Others");
            FirstFind = RangeFind;
            while (RangeFind != null && (FirstFind.Row != RangeFind.Row || bFirst))
            {
                bFirst = false;
                this.Cells[RangeFind.Row, 3].Value = "xxx";

                RangeFind = this.Cells.Find(What: "Total Others", After: RangeFind, SearchDirection: Excel.XlSearchDirection.xlNext);
            }

            return true;
        }

        public bool SetupHeaders()
        {
            CAE___Grandes_Actores___Grandes_Cadenas.Data1 oData1 = Globals.Data1;
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
                HeaderValue = oData1.Range[oPair.Key].Value;
                System.Text.RegularExpressions.Match oMath = oRegexMat.Match(HeaderValue);
                if (oMath.Success)
                {
                    String CurrentTittle = this.Range[oPair.Value].Value;
                    CurrentTittle = CurrentTittle.Replace("YYYY", oMath.Value.Substring(1));
                    this.Range[oPair.Value].Value = CurrentTittle;
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
                HeaderValue = oData1.Range[oPair.Key].Value;
                System.Text.RegularExpressions.Match oMathMonth = oRegexMatMonth.Match(HeaderValue);
                if (oMathMonth.Success)
                {
                    String sMonth = oMathMonth.Value.Substring(1, oMathMonth.Value.Substring(1).IndexOf('/'));
                    String sYear = oMathMonth.Value.Substring(oMathMonth.Value.Length - 4, 4);
                    DateTime oDate = new DateTime(int.Parse(sYear), int.Parse(sMonth), 1);
                    this.Range[oPair.Value].Value = oDate;
                }
            }

            return true;
        }

        public bool paint_data()
        {
            //IMSClasses.DBHelper.db oDB = Globals.ThisWorkbook.oDb;
            //DataSet oDataSet = oDB.ExecuteQuery("select dbo.[fn_CountSpaces](col1) lvl, * from dbo.CAE_GAFP_Data_2  where dbo.[fn_CountSpaces](col1)=2", CommandType.Text);

            // 1 get de data sheets reference
            CAE___Grandes_Actores___Grandes_Cadenas.Data1 oData1 = Globals.Data1;

            // 2 set hierarchy data
            Excel.Range oRangeStart = oData1.Range["A2:B2"];
            Excel.Range oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            Excel.Range oRangeColumn = oData1.Range[oRangeStart, oRangeEnd];

            oRangeColumn.Copy(this.Range["B4"]);
            //setup max length
            int iMaxRow = 0;
            iMaxRow = oRangeEnd.Row;


            //copy lvl
            oRangeStart = oData1.Range["BC2"];//--> this is lvl
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["A4"]);

            /* 
                AW --> D "Sales Euros PUB YTD/5/2014  (Thousands)"
                AX --> E "Sales Euros PUB YTD/5/2014 %PPG Previous Year (Absolute)"	
                AY --> F "Sales Euros PUB YTD/5/2014 %V (Absolute)"
                AW/W   --> R precio_medio?
            */
            String sFormula = "=SI(Data1!W2>0;Data1!AW2/Data1!W2;\"--\")";
            //oRangeStart = oData1.Range["W2:Y2"];
            //oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range["W2:Y" + iMaxRow.ToString()];
            oRangeColumn.Copy(this.Range["D4:F4"]);
            this.Range["G4"].FormulaLocal = sFormula;
            oRangeColumn = this.Range["G4:G"  + (iMaxRow+2).ToString()]; //+ oRangeEnd.Row.ToString()];
            this.Range["G4"].AutoFill(oRangeColumn);



            /*
            AZ --> H
            BA --> I
            BB --> J
            AZ/Z   --> K precio_medio?
            */
            sFormula = "=SI(Data1!Z2>0;Data1!AZ2/Data1!Z2;\"--\")";
            //oRangeStart = oData1.Range["Z2:AB2"];
            //oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range["Z2:AB" + iMaxRow.ToString()];
            oRangeColumn.Copy(this.Range["H4:J4"]);
            this.Range["K4"].FormulaLocal = sFormula;
            oRangeColumn = this.Range["K4:K"  + (iMaxRow+2).ToString()]; //+ oRangeEnd.Row.ToString()];
            this.Range["K4"].AutoFill(oRangeColumn);



            /* may-2014
            AC --> L
            AD --> M
            AC/C   --> N precio_medio?
            */
            //oRangeStart = oData1.Range["C2:D2"];
            //oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range["C2:D" + iMaxRow.ToString()];
            oRangeColumn.Copy(this.Range["L4:M4"]);
            this.Range["N4"].FormulaLocal = "=SI(Data1!C2>0;Data1!AC2/Data1!C2;\"--\")"; //"='Data1'!AC2/'Data1'!C2";
            oRangeColumn = this.Range["N4:N"  + (iMaxRow+2).ToString()]; //+ oRangeEnd.Row.ToString()];
            this.Range["N4"].AutoFill(oRangeColumn);

            /*
            mar-15
            AE --> O
            AF --> P
            AG --> Q
            AE/E   --> R precio_medio?
            */
            //oRangeStart = oData1.Range["E2:G2"];
            //oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range["E2:G" + iMaxRow.ToString()];
            oRangeColumn.Copy(this.Range["O4:Q4"]);
            this.Range["R4"].FormulaLocal = "=SI(Data1!E2>0;Data1!AE2/Data1!E2;\"--\")"; //"='Data1'!AE2/'Data1'!E2";
            oRangeColumn = this.Range["R4:R"  + (iMaxRow+2).ToString()]; //+ oRangeEnd.Row.ToString()];
            this.Range["R4"].AutoFill(oRangeColumn);

            /*
            abr-2015
            AH --> S
            AI --> T
            AJ --> U
            AH/H   --> V precio_medio?
            */
            //oRangeStart = oData1.Range["H2:J2"];
            //oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range["H2:J" + iMaxRow.ToString()];
            oRangeColumn.Copy(this.Range["S4:U4"]);
            this.Range["V4"].FormulaLocal = "=SI(Data1!H2>0;Data1!AH2/Data1!H2;\"--\")"; //"='Data1'!AH2/'Data1'!H2";
            oRangeColumn = this.Range["V4:V"  + (iMaxRow+2).ToString()]; //+ oRangeEnd.Row.ToString()];
            this.Range["V4"].AutoFill(oRangeColumn);

            /*
            may-15 
            AK --> W
            AL --> X
            AM --> Y
            AK/K   --> Z
            */
            //oRangeStart = oData1.Range["K2:M2"];
            //oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range["K2:M" + iMaxRow.ToString()];
            oRangeColumn.Copy(this.Range["W4:Y4"]);
            this.Range["Z4"].FormulaLocal = "=SI(Data1!K2>0;Data1!AK2/Data1!K2;\"--\")"; //"='Data1'!AK2/'Data1'!K2";
            oRangeColumn = this.Range["Z4:Z"  + (iMaxRow+2).ToString()]; //+ oRangeEnd.Row.ToString()];
            this.Range["Z4"].AutoFill(oRangeColumn);

            /*
            Derniers 3 mois
            AT --> AA
            AU --> AB
            AV --> AC
            AT/T   --> AD Precio Medio
            */
            //oRangeStart = oData1.Range["T2:V2"];
            //oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range["T2:V" + iMaxRow.ToString()];
            oRangeColumn.Copy(this.Range["AA4:AC4"]);
            this.Range["AD4"].FormulaLocal = "=SI(Data1!T2>0;Data1!AT2/Data1!T2;\"--\")"; //"='Data1'!AT2/'Data1'!T2";
            oRangeColumn = this.Range["AD4:AD"  + (iMaxRow+2).ToString()]; //+ oRangeEnd.Row.ToString()];
            this.Range["AD4"].AutoFill(oRangeColumn);

            /*
            Año 2013
            AN --> AE
            AO --> AF
            AP --> AG
            AN/N   --> AH Precio Medio
            */
            //oRangeStart = oData1.Range["N2:P2"];
            //oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range["N2:P" + iMaxRow.ToString()];
            oRangeColumn.Copy(this.Range["AE4:AG4"]);
            this.Range["AH4"].FormulaLocal = "=SI(Data1!N2>0;Data1!AN2/Data1!N2;\"--\")"; ////"='Data1'!AN2/'Data1'!N2";
            oRangeColumn = this.Range["AH4:AH"  + (iMaxRow+2).ToString()]; //+ oRangeEnd.Row.ToString()];
            this.Range["AH4"].AutoFill(oRangeColumn);

            /*
            Año 2014
            AQ --> AI
            AR --> AJ
            AS --> AK
            AQ/Q   --> AL Precio Medio
             
             */
            //oRangeStart = oData1.Range["Q2:S2"];
            //oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData1.Range["Q2:S" + iMaxRow.ToString()];
            oRangeColumn.Copy(this.Range["AI4:AK4"]);
            this.Range["AL4"].FormulaLocal = "=SI(Data1!Q2>0;Data1!AQ2/Data1!Q2;\"--\")"; //"='Data1'!AQ2/'Data1'!Q2";
            oRangeColumn = this.Range["AL4:AL"  + (iMaxRow+2).ToString()]; //+ oRangeEnd.Row.ToString()];
            this.Range["AL4"].AutoFill(oRangeColumn);

            //d - al
            List<Excel.Range> oRangesToFillEmptys = new List<Excel.Range>();
            oRangesToFillEmptys.Add(this.Range["D3:AL" + (iMaxRow + 2).ToString()]);
            Helppers.fillEmptySpaces(this, (iMaxRow + 2), oRangesToFillEmptys);


            return true;
        }

        private void Hoja3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Hoja3_Startup);
            this.Shutdown += new System.EventHandler(Hoja3_Shutdown);
        }

        #endregion

    }
}
