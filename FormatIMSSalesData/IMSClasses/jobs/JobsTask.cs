using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IMSClasses.Jobs
{
    public class JobsTask
    {
        public List<Task> Tasks;

        public JobsTask()
        {
            this.Tasks = new List<Task>();
        }
        public bool addJob (Task oTask)
        {
            this.Tasks.Add(oTask);
            return true;
        }
    }
}
