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

namespace TestTemplate
{
    public partial class ReportTest
    {
        private void Hoja2_Startup(object sender, System.EventArgs e)
        {
            this.paint_data();
            this.format_data();

            
        }

        private void Hoja2_Shutdown(object sender, System.EventArgs e)
        {
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
                if (oCell.Value == 6 || oCell.Value == 1 || oCell.Value == 2 || oCell.Value == 7 || oCell.Value == 5 || oCell.Value == 12)
                {
                    oRangeStart = this.Range["B" + oCell.Row.ToString()];
                    oRangeEnd = oRangeStart.End[Excel.XlDirection.xlToRight];
                    oRangeColumn = this.Range[oRangeStart, oRangeEnd];
                }

                if (oCell.Value == 6 || oCell.Value == 1)
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

                if (oCell.Value == 7 || oCell.Value == 2)
                {
                    oRangeColumn.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    oRangeColumn.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    oRangeColumn.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5;
                    oRangeColumn.Interior.TintAndShade = 0.799981688894314;
                    oRangeColumn.Interior.PatternTintAndShade = 0;
                }

                if (oCell.Value == 12)
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
            ArrayList oRangeToAdd = new ArrayList();

            Excel.Range FirstGroupRange = null, LastGroupRange = null, PreviousParentRange = null, PreviousCell = null;

            oRangeStart = this.Range["A4"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = this.Range[oRangeStart, oRangeEnd];
            Excel.Range AllCellsRange = this.Range[oRangeStart, oRangeEnd];

            PreviousParentRange = null;
            oRangeGroups.AddRange(Helppers.getGroupRanges(AllCellsRange, this));
        
            //agrupar
            foreach (Excel.Range oRangeToGroup in oRangeGroups)
            {
                oRangeToGroup.Rows.Group();
            }

            /*
            oRangeStart = this.Range["A4"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = this.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Value = "";*/

            return true;
        }

        public bool paint_data()
        {
            //IMSClasses.DBHelper.db oDB = Globals.ThisWorkbook.oDb;
            //DataSet oDataSet = oDB.ExecuteQuery("select dbo.[fn_CountSpaces](col1) lvl, * from dbo.CAE_GAFP_Data_2  where dbo.[fn_CountSpaces](col1)=2", CommandType.Text);

            // 1 get de data sheets reference
            TestTemplate.Data oData = Globals.Data;


            // 2 set hierarchy data
            Excel.Range oRangeStart = oData.Range["A2:B2"];
            Excel.Range oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            Excel.Range oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["B4"]);

            //copy lvl
            oRangeStart = oData.Range["BC2"];//--> this is lvl
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["A4"]);

            /* 
                AW --> D "Sales Euros PUB YTD/5/2014  (Thousands)"
                AX --> E "Sales Euros PUB YTD/5/2014 %PPG Previous Year (Absolute)"	
                AY --> F "Sales Euros PUB YTD/5/2014 %V (Absolute)"
                AW/W   --> R precio_medio?
            */
            String sFormula = "=SI(Data!W2>0;Data!AW2/Data!W2;\"--\")";

            oRangeStart = oData.Range["AW2:AY2"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["D4:F4"]);
            this.Range["G4"].FormulaLocal = sFormula;
            oRangeColumn = this.Range["G4:G" + oRangeEnd.Row.ToString()];
            this.Range["G4"].AutoFill(oRangeColumn);


            /*
            AZ --> H
            BA --> I
            BB --> J
            AZ/Z   --> K precio_medio?
            */
            sFormula = "=SI(Data!Z2>0;Data!AZ2/Data!Z2;\"--\")";
            oRangeStart = oData.Range["AZ2:BB2"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["H4:J4"]);
            this.Range["K4"].FormulaLocal = sFormula;
            oRangeColumn = this.Range["K4:K" + oRangeEnd.Row.ToString()];
            this.Range["K4"].AutoFill(oRangeColumn);

            /* may-2014
            AC --> L
            AD --> M
            AC/C   --> N precio_medio?
            */
            oRangeStart = oData.Range["AC2:AD2"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["L4:M4"]);
            this.Range["N4"].FormulaLocal = "=SI(Data!C2>0;Data!AC2/Data!C2;\"--\")"; //"='Data'!AC2/'Data'!C2";
            oRangeColumn = this.Range["N4:N" + oRangeEnd.Row.ToString()];
            this.Range["N4"].AutoFill(oRangeColumn);


            /*
            mar-15
            AE --> O
            AF --> P
            AG --> Q
            AE/E   --> R precio_medio?
            */
            oRangeStart = oData.Range["AE2:AG2"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["O4:Q4"]);
            this.Range["R4"].FormulaLocal = "=SI(Data!E2>0;Data!AE2/Data!E2;\"--\")"; //"='Data'!AE2/'Data'!E2";
            oRangeColumn = this.Range["R4:R" + oRangeEnd.Row.ToString()];
            this.Range["R4"].AutoFill(oRangeColumn);


            /*
            abr-2015
            AH --> S
            AI --> T
            AJ --> U
            AH/H   --> V precio_medio?
            */
            oRangeStart = oData.Range["AH2:AJ2"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["S4:U4"]);
            this.Range["V4"].FormulaLocal = "=SI(Data!H2>0;Data!AH2/Data!H2;\"--\")"; //"='Data'!AH2/'Data'!H2";
            oRangeColumn = this.Range["V4:V" + oRangeEnd.Row.ToString()];
            this.Range["V4"].AutoFill(oRangeColumn);

            /*
            may-15 
            AK --> W
            AL --> X
            AM --> Y
            AK/K   --> Z
            */
            oRangeStart = oData.Range["AK2:AM2"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["W4:Y4"]);
            this.Range["Z4"].FormulaLocal = "=SI(Data!K2>0;Data!AK2/Data!K2;\"--\")"; //"='Data'!AK2/'Data'!K2";
            oRangeColumn = this.Range["Z4:Z" + oRangeEnd.Row.ToString()];
            this.Range["Z4"].AutoFill(oRangeColumn);

            /*
            Derniers 3 mois
            AT --> AA
            AU --> AB
            AV --> AC
            AT/T   --> AD Precio Medio
            */
            oRangeStart = oData.Range["AT2:AV2"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["AA4:AC4"]);
            this.Range["AD4"].FormulaLocal = "=SI(Data!T2>0;Data!AT2/Data!T2;\"--\")"; //"='Data'!AT2/'Data'!T2";
            oRangeColumn = this.Range["AD4:AD" + oRangeEnd.Row.ToString()];
            this.Range["AD4"].AutoFill(oRangeColumn);

            /*
            Año 2013
            AN --> AE
            AO --> AF
            AP --> AG
            AN/N   --> AH Precio Medio
            */
            oRangeStart = oData.Range["AN2:AP2"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["AE4:AG4"]);
            this.Range["AH4"].FormulaLocal = "=SI(Data!N2>0;Data!AN2/Data!N2;\"--\")"; ////"='Data'!AN2/'Data'!N2";
            oRangeColumn = this.Range["AH4:AH" + oRangeEnd.Row.ToString()];
            this.Range["AH4"].AutoFill(oRangeColumn);


            /*
            Año 2014
            AQ --> AI
            AR --> AJ
            AS --> AK
            AQ/Q   --> AL Precio Medio
             
             */
            oRangeStart = oData.Range["AQ2:AS2"];
            oRangeEnd = oRangeStart.End[Excel.XlDirection.xlDown];
            oRangeColumn = oData.Range[oRangeStart, oRangeEnd];
            oRangeColumn.Copy(this.Range["AI4:AK4"]);
            this.Range["AL4"].FormulaLocal = "=SI(Data!Q2>0;Data!AQ2/Data!Q2;\"--\")"; //"='Data'!AQ2/'Data'!Q2";
            oRangeColumn = this.Range["AL4:AL" + oRangeEnd.Row.ToString()];
            this.Range["AL4"].AutoFill(oRangeColumn);


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
