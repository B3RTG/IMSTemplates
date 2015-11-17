using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IMSReportServices.Models
{
    public class ConfigurationModel
    {
        public String ConnectionString { get; set; }
        public String UploadPath;
        public RabbittConfig RabbittMQ;
        public class RabbittConfig
        {
            public String ImportQueueName;
            public String ExcelQueueName;
            public String Server;
        }

        public String TaskManager_ProccesFolder;
        public String TaskManager_ProccesFolder_Out;

        public ConfigurationModel()
        {
            this.ConnectionString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;
            this.UploadPath = System.Web.Configuration.WebConfigurationManager.AppSettings["TaskManager_UploadPath"].ToString();

            this.RabbittMQ = new RabbittConfig();
            this.RabbittMQ.ImportQueueName = System.Web.Configuration.WebConfigurationManager.AppSettings["TaskManager_RabbitMQ_Import_Queue"].ToString();
            this.RabbittMQ.ExcelQueueName = System.Web.Configuration.WebConfigurationManager.AppSettings["TaskManager_RabbitMQ_Excel_Queue"].ToString();
            this.RabbittMQ.Server = System.Web.Configuration.WebConfigurationManager.AppSettings["TaskManager_RabbitMQ_Server"].ToString();

            this.TaskManager_ProccesFolder = System.Web.Configuration.WebConfigurationManager.AppSettings["TaskManager_ProccesFolder"].ToString();
            this.TaskManager_ProccesFolder_Out = System.Web.Configuration.WebConfigurationManager.AppSettings["TaskManager_ProccesFolder_Out"].ToString();
        }
    }
}