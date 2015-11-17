using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMSClasses
{
    public class ConfigurationHelpper
    {
        public String ConnectionString;
        public String ConnectionStringSQL;

        public int JobID;
        public String OutputPath;
        public String ProccesPath;


        public ConfigurationHelpper()
        {
            this.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DBConnection"].ToString();
            this.ConnectionStringSQL = System.Configuration.ConfigurationManager.ConnectionStrings["DBConnectionSQLCli"].ToString();

            this.ProccesPath = System.Configuration.ConfigurationManager.AppSettings["ProccesFolder"].ToString();
            this.OutputPath = System.Configuration.ConfigurationManager.AppSettings["OutFolder"].ToString();
            this.JobID = int.Parse(System.Configuration.ConfigurationManager.AppSettings["JobID"].ToString());
        }
    }
}
