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

using IMSClasses.Jobs;
using IMSClasses.DBHelper;
using IMSClasses;

namespace CAE___Grandes_Actores_Parafarmacias
{
    public partial class ThisWorkbook
    {
        public ConfigurationHelpper oCfg;
        public String StatusMessage;
        public bool StatusCorrect;

        public IMSClasses.Jobs.Job oJob;
        public IMSClasses.DBHelper.db oDb;

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            oCfg = new ConfigurationHelpper();
            oDb = new db(oCfg.ConnectionStringSQL.ToString());


            //get the job data
            oJob = IMSClasses.Jobs.Job.getInstance(oDb.getJob(oCfg.JobID)["JSON"].ToString());

            this.RemoveCustomization();
            this.StatusMessage = "";
            this.StatusCorrect = true;
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
