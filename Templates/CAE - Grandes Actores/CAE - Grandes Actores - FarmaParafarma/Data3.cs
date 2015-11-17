using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace CAE___Grandes_Actores___FarmaParafarma
{
    public partial class Data3
    {
        public Excel.QueryTable oTable;
        private void Hoja3_Startup(object sender, System.EventArgs e)
        {
            IMSClasses.ConfigurationHelpper oCfg = Globals.ThisWorkbook.oCfg;
            IMSClasses.Jobs.Job oJob = Globals.ThisWorkbook.oJob;
            IMSClasses.DBHelper.db oDB = Globals.ThisWorkbook.oDb;

            try
            {
                oTable = Helppers.importData(oCfg.ConnectionString, "SELECT * FROM " + oJob.SQLParameters.TableName.Replace(@"%identity%", "3"), this);
            }
            catch
            {
                this.oTable = null;
                Globals.ThisWorkbook.StatusMessage = "Error getting data in sheet 3";
                Globals.ThisWorkbook.StatusCorrect = false;


                oJob.ReportStatus.Message = Globals.ThisWorkbook.StatusMessage;
                oJob.ReportStatus.Status = "ERRO";
                oDB.updateJob(oJob.Serialize(), oJob.JOBID);
            }


            if (!Globals.ThisWorkbook.StatusCorrect || this.oTable == null)
            {
                this.oTable = null;
                Globals.ThisWorkbook.StatusMessage = "Error getting data in sheet 3";
                Globals.ThisWorkbook.StatusCorrect = false;

                oJob.ReportStatus.Message = Globals.ThisWorkbook.StatusMessage;
                oJob.ReportStatus.Status = "ERRO";
                oDB.updateJob(oJob.Serialize(), oJob.JOBID);
            }
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
