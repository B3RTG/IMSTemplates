using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

using IMSClasses.Jobs;
using IMSClasses.DBHelper;

namespace IMSReportServices.Controllers
{
    public class JobController : ApiController
    {
        IMSClasses.DBHelper.db oDB;
        Models.ConfigurationModel oConfig = new Models.ConfigurationModel();
        // GET api/job
        public IEnumerable<Job> Get()
        {
            List<Job> oJobList = new List<Job>();
            this.oDB = new db(oConfig.ConnectionString);
            System.Data.DataTable oJobsTable = oDB.getConfiguration();

            foreach(System.Data.DataRow oJobRow in oJobsTable.Rows)
            {
                IMSClasses.Jobs.Job oJobToAdd = Job.getInstance(oJobRow["JSON"].ToString());
                oJobToAdd.CurrentTaskStatus = oJobRow["CurrentTaskStatus"].ToString();
                oJobList.Add(oJobToAdd);
            }

            return oJobList;
        }

        // GET api/job/5
        public Job Get(Int64 id)
        {
            this.oDB = new db(oConfig.ConnectionString);
            System.Data.DataRow oJobRow = oDB.getJob(id);

            return Job.getInstance(oJobRow["JSON"].ToString());
        }

        // POST api/job
        public void Post([FromBody]string value)
        {
        }

        // PUT api/job/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/job/5
        public void Delete(int id)
        {
        }
    }
}
