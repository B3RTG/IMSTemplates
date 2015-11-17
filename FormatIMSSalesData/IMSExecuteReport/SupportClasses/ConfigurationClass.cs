using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace RabbittMQ_ExcelFormat
{
    public class ConfigurationClass
    {

        public class RabbittConfig
        {
            public String ImportQueueName;
            public String ExcelQueueName;
            public String Server;
        }

        public String ConnectionString;
        public RabbittConfig RabbittMQ;

        /*
        <add key="ProccesFolder" value="C:\Dev\IMS\@PROC\data"/>
        <add key="ProccesFolder_in" value="in"/>
        <add key="ProccesFolder_Out" value="out"/>
        <add key="Template" value="Template"/>
        */
        public class PathsClass
        {
            public String MainFolder;
            public String InFolder;
            public String OutFolder;
            public String TemplateFolder;
        }

        public PathsClass Paths;

        public ConfigurationClass()
        {
            //this.ConnectionString = ConfigurationManager.ConnectionStrings["DBServer"].ToString();


            this.ConnectionString = ConfigurationManager.ConnectionStrings["dbServer"].ToString();
            this.RabbittMQ = new RabbittConfig();
            this.RabbittMQ.ImportQueueName = ConfigurationManager.AppSettings["RabbitMQ_Import_Queue"].ToString();
            this.RabbittMQ.ExcelQueueName = ConfigurationManager.AppSettings["RabbitMQ_Excel_Queue"].ToString();
            this.RabbittMQ.Server = ConfigurationManager.AppSettings["RabbitMQ_Server"].ToString();


            this.Paths = new PathsClass();
            this.Paths.MainFolder = ConfigurationManager.AppSettings["ProccesFolder"].ToString();
            this.Paths.InFolder = ConfigurationManager.AppSettings["ProccesFolder_in"].ToString();
            this.Paths.OutFolder = ConfigurationManager.AppSettings["ProccesFolder_Out"].ToString();
            this.Paths.TemplateFolder = ConfigurationManager.AppSettings["Template"].ToString();
        }
    }
}
