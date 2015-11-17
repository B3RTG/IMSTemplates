using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Top50Farmacias
{
    public class ConfigurationHelpper
    {
        public String ConnectionString;
        public String ConnectionStringSQL;
        //public String SQLQuery;
        //public String Output_ExcelFilename;
        public String OutputPath;
        public String ProccesPath;

        public int JobID;

        public ConfigurationHelpper()
        {
            this.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DBConnection"].ToString();
            this.ConnectionStringSQL = System.Configuration.ConfigurationManager.ConnectionStrings["DBConnectionSQLCli"].ToString();
            this.JobID = int.Parse(System.Configuration.ConfigurationManager.AppSettings["JobID"].ToString());

            /*
             <add key="ProccesFolder" value="C:\Dev\IMS\@PROC\data"/>
             <add key="OutFolder" value="out"/>
             */
            this.ProccesPath = System.Configuration.ConfigurationManager.AppSettings["ProccesFolder"].ToString();
            this.OutputPath = System.Configuration.ConfigurationManager.AppSettings["OutFolder"].ToString();

            //this.SQLQuery = System.Configuration.ConfigurationManager.AppSettings["SQLQuery"].ToString();
            //this.Output_ExcelFilename = System.Configuration.ConfigurationManager.AppSettings["Output_Excel_Filename"].ToString();
            
        }
    }
}
