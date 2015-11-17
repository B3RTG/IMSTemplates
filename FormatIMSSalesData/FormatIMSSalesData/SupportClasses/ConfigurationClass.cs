using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace FormatIMSSalesData
{
    public class ConfigurationClass
    {

        public class RabbittConfig
        {
            public String ImportQueueName;
            public String ExcelQueueName;
            public String Server;
        }

        
        public RabbittConfig RabbittMQ;

        public String ConnectionString;
        public String LogFilePath;

        public String LocalOutPath;
        public String LocalInPath;
        public Boolean Recursive;
        public String ExtensionFilter;  
  



        public ConfigurationClass()
        {
            //this.ConnectionString = ConfigurationManager.ConnectionStrings["DBServer"].ToString();
            this.LogFilePath = ConfigurationManager.AppSettings["LogFilePath"].ToString();
            this.LogFilePath = this.LogFilePath.Replace("%%date%%", DateTime.Now.ToString("yyyyMMdd_hhmmss"));

            this.ConnectionString = ConfigurationManager.AppSettings["db_ConnectionString"].ToString();

            this.LocalInPath = ConfigurationManager.AppSettings["LocalPath"].ToString();
            //this.LocalOutPath = ConfigurationManager.AppSettings["LocalOutPath"].ToString();
            this.Recursive = Convert.ToBoolean(ConfigurationManager.AppSettings["Recursive"].ToString());
            this.ExtensionFilter = ConfigurationManager.AppSettings["ExtensionFilter"].ToString();

            this.RabbittMQ = new RabbittConfig();
            this.RabbittMQ.ImportQueueName = ConfigurationManager.AppSettings["RabbitMQ_Import_Queue"].ToString();
            this.RabbittMQ.ExcelQueueName = ConfigurationManager.AppSettings["RabbitMQ_Excel_Queue"].ToString();
            this.RabbittMQ.Server = ConfigurationManager.AppSettings["RabbitMQ_Server"].ToString();

        }
    }
}
