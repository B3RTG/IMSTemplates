using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

using IMSClasses.Jobs;

namespace IMSReportServices.Controllers
{
    public class TaskController : ApiController
    {
        IMSClasses.DBHelper.db oDB;
        Models.ConfigurationModel oConfig = new Models.ConfigurationModel();

        // GET api/task
        public IEnumerable<Task> Get()
        {
            List<Task> oTaskList = new List<Task>();
            this.oDB = new IMSClasses.DBHelper.db(oConfig.ConnectionString);
            System.Data.DataTable oJobsTable = oDB.getTaskList();

            foreach (System.Data.DataRow oJobRow in oJobsTable.Rows)
            {
                Task oTaskToAdd = Task.getInstance(oJobRow["JSON"].ToString());
                oTaskList.Add(oTaskToAdd);
            }



            return oTaskList;
        }

        // GET api/task/5
        public Task Get(int id)
        {
            this.oDB = new IMSClasses.DBHelper.db(oConfig.ConnectionString);
            System.Data.DataRow oJobRow = oDB.getTask(id);

            return Task.getInstance(oJobRow["JSON"].ToString());
        }

        // POST api/task
        public void Post([FromBody]string value)
        {

        }

        // PUT api/task/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/task/5
        public void Delete(int id)
        {
        }
    }
}
