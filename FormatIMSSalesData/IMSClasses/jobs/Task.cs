using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace IMSClasses.Jobs
{
    public class Task
    {
        public Job oJob;
        public Int64 TaskID;
        public DateTime? CreateDate;
        public DateTime? UpdateDate;
        public String StatusCurrent;
        public String StatusFinal;
        public String TaskComments;

        public Task()
        {        }

        public Task(Int64 iTaskID, DateTime? CreateDate, DateTime? UpdateDate, String CurrentStatus, String FinalStatus, String JobJSON)
        {
            this.TaskID = iTaskID;
            this.CreateDate = CreateDate;
            this.UpdateDate = UpdateDate;
            this.StatusCurrent = CurrentStatus;
            this.StatusFinal = FinalStatus;
            this.oJob = Job.getInstance(JobJSON);

        }

        public static Task getInstance(String sJSON)
        {
            JavaScriptSerializer oSerializaer = new JavaScriptSerializer();
            Task oTask = oSerializaer.Deserialize<Task>(sJSON);

            return oTask;
        }

        public String Serialize()
        {
            JavaScriptSerializer oSerializaer = new JavaScriptSerializer();
            return oSerializaer.Serialize(this);
        }

    }
}
