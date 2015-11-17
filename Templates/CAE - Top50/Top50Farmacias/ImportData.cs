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

using System.Data;
using System.Data.SqlClient;

namespace Top50Farmacias
{
    public partial class ImportData
    {
        public Excel.QueryTable oTable;
        private void Hoja1_Startup(object sender, System.EventArgs e)
        {
            ConfigurationHelpper oCfg = Globals.ThisWorkbook.oCfg;
            IMSClasses.Jobs.Job oJob = Globals.ThisWorkbook.oJob;
            IMSClasses.DBHelper.db oDB = Globals.ThisWorkbook.oDb;

            

            try
            {
                String sSqlQuery = "select * from " + oJob.SQLParameters.TableName;
                if (Globals.ThisWorkbook.StatusCorrect && !this.importData(oCfg.ConnectionString, sSqlQuery))
                { //error obteniendo datos
                    Globals.ThisWorkbook.StatusMessage = "Error getting data";
                    Globals.ThisWorkbook.StatusCorrect = false;

                    oJob.ReportStatus.Message = "Error getting data";
                    oJob.ReportStatus.Status = "ERRO";
                    oJob.ReportStatus.ExecutionDate = DateTime.Now;
                    oDB.updateJob(oJob.Serialize(), oJob.JOBID);

                    //Globals.ThisWorkbook.Close(0);
                }
            }
            catch (Exception eImportData)
            {
                Globals.ThisWorkbook.StatusCorrect = false;
                Globals.ThisWorkbook.StatusMessage = "Error(Not Controled)--> Sheet Report --> Exception Message: " + eImportData.Message.ToString();

                oJob.ReportStatus.Message = Globals.ThisWorkbook.StatusMessage;
                oJob.ReportStatus.Status = "ERRO";
                oDB.updateJob(oJob.Serialize(), oJob.JOBID);

                //Globals.ThisWorkbook.Close(0);
            }
            
            
           
        }

        public bool importData(String sConnectionString, String sSqlQuery)
        {
            Boolean bDone = false;
            try
            {
                Excel.Range oRange = this.Range["A1"];
                this.oTable = this.QueryTables.Add(sConnectionString, oRange);
                this.oTable.CommandType = Excel.XlCmdType.xlCmdSql;
                this.oTable.CommandText = sSqlQuery;
                this.oTable.Refresh();
                bDone = true;

                for (int i = 1; i <= Globals.ThisWorkbook.Connections.Count; i++)
                    Globals.ThisWorkbook.Connections[i].Delete();
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
                bDone   = false;
            }

            return bDone;
        }

        private void Hoja1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Hoja1_Startup);
            this.Shutdown += new System.EventHandler(Hoja1_Shutdown);
            //this.EndInit += new System.EventHandler(Hoja1_End);
        }

        #endregion

    }
}
