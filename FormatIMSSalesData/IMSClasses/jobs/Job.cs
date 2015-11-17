using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace IMSClasses.Jobs
{
    public class Job
    {

        public Int64 JOBID { get; set; }
        public String JOBCODE { get; set; }
        public String Tittle { get; set; }
        public IMSClasses.Jobs.InputParameters InputParameters;
        public IMSClasses.Jobs.OutputParameters OutputParameters;
        public IMSClasses.Jobs.SQLParameters SQLParameters;
        public IMSClasses.Jobs.Schelude Schelude;
        public IMSClasses.Jobs.ExecutionStatus ImportStatus;
        public IMSClasses.Jobs.ExecutionStatus ReportStatus;
        public IMSClasses.Jobs.ExecutionStatus SendStatus;

        public String CurrentTaskStatus;

        public String PluginName;
        
        public Job( )
        {            
        }

        /***
         * this is for test purpose only
         * */
        public Job(String sTittle)
        {
            this.JOBCODE = "";
            this.Tittle = "";
            this.InputParameters = new InputParameters();
            this.InputParameters.Files.Add(new File("name", "path"));
            this.OutputParameters = new OutputParameters();
            this.OutputParameters.OriginalFile = new File("name", "path");
            this.OutputParameters.DestinationFile = new File("name", "path");
            this.OutputParameters.channel = "FILESYSTEM";
            this.OutputParameters.MailParameters = new OutputParameters.mailparameters();
            this.Schelude = new Schelude();
            this.SQLParameters = new SQLParameters();
            this.SQLParameters.TableName = "tablename";
            this.Schelude.Periodicity = "Monthly";
            this.Schelude.Day = 3;
            this.Schelude.WeekDays = "Lunes,Martes";
            this.Schelude.Hour = 14;
            this.Schelude.Minut = 14;
            

        }

        public static Job getInstance(String sJSON)
        {
            JavaScriptSerializer oSerializaer = new JavaScriptSerializer();
            Job oJob = oSerializaer.Deserialize<Job>(sJSON);

            return oJob;
        }

        public String Serialize()
        {
            JavaScriptSerializer oSerializaer = new JavaScriptSerializer();
            return oSerializaer.Serialize(this);
        }
    }
}
