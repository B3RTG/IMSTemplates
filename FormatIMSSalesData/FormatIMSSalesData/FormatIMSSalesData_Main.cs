using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IMSClasses;

//using IMSClasses.excellHelpper;
using IMSClasses.RabbitMQ;


namespace FormatIMSSalesData
{
    class FormatIMSSalesData_Main
    {
        static void Main(string[] args)
        {

            //IMSClasses.excellHelpper.ExcelHelpper.executeTest(@"C:\Dev\IMS\testbook.xlsm");


            ConfigurationClass oConf = new ConfigurationClass();
            IMSClasses.Jobs.JobsTask oJobsTask = new IMSClasses.Jobs.JobsTask();

            IMSClasses.DBHelper.db _db = new IMSClasses.DBHelper.db(oConf.ConnectionString);
            //System.Data.DataTable oDataTable = _db.getConfiguration();
            System.Data.DataTable oDataTable = _db.getPendingTask();

            // Load Task from DB. Maibe we have to put here the schelude logic.
            foreach (System.Data.DataRow oRow in oDataTable.Rows)
            {
                //IMSClasses.Jobs.Job oJob = IMSClasses.Jobs.Job.getInstance(oRow["JSON"].ToString());
                //prepare task to enqueue
                IMSClasses.Jobs.Task oTask = IMSClasses.Jobs.Task.getInstance(oRow["JSON"].ToString());
                if (oTask.oJob.JOBCODE == null) oTask.oJob = IMSClasses.Jobs.Job.getInstance(oRow["JobJSON"].ToString());
                
                oTask.TaskID = Int64.Parse(oRow["PK_TASK"].ToString());
                _db.updateTask(oTask);
                
                if (oTask.oJob.JOBID == 0)
                {
                    oTask.oJob.JOBID = Int64.Parse(oRow["FK_JOB_ID"].ToString());
                    _db.updateJob(oTask.oJob.Serialize(), oTask.oJob.JOBID);
                }
                oJobsTask.addJob(oTask);
            }

            IMSClasses.RabbitMQ.MessageQueue oImportQueue = new MessageQueue(oConf.RabbittMQ.Server, null, oConf.RabbittMQ.ImportQueueName);

            foreach (IMSClasses.Jobs.Task oTaskToProc in oJobsTask.Tasks)
            {
                //mirar si se tiene que procesar, si es asi, lo añadimos a la cola de importaciones.
                //if(oJobToProc.JOBID==3)
                if(true)
                {//stuff here
                    oTaskToProc.StatusCurrent = "IMQU";
                    oTaskToProc.UpdateDate = DateTime.Now;
                    oTaskToProc.TaskComments = "QUEUED for import.";
                    

                    _db.updateTask(oTaskToProc);
                    //oImportQueue.addMessage(int.Parse(oTaskToProc.JOBID.ToString()));
                    oImportQueue.addMessage(int.Parse(oTaskToProc.TaskID.ToString()));
                }
                else
                {//do nothing

                }
            }

            oImportQueue.close();

            /*
            foreach (IMSClasses.Jobs.Job oJobToProc in oJobsTask.Jobs)
            {
                bool bCorrect = true;
                String sError = "";
                System.Data.DataTable dt = null;
                
                try
                {
                    dt = ExcelHelpper.getExcelData(oJobToProc.InputParameters.Files[0].FileName, oJobToProc.SQLParameters.TableName);
                }
                catch(Exception xlException)
                {
                    bCorrect = false;
                    sError = "Error getting excel data --> Exception --> " + xlException.Message;
                }
                
                if(bCorrect)
                {
                    try
                    {
                        bCorrect = _db.LoadTable(dt);
                    }
                    catch(Exception dbException)
                    {
                        bCorrect = false;
                        sError = "Error loading data in DB --> Exception --> " + dbException.Message;
                    }
                    
                } 

                if(!bCorrect)
                { //failure update
                    oJobToProc.ImportStatus.ExecutionDate = DateTime.Now;
                    oJobToProc.ImportStatus.Message = sError;
                    oJobToProc.ImportStatus.Status = "ERRO";
                }
                else
                { //correct job update
                    oJobToProc.ImportStatus.ExecutionDate = DateTime.Now;
                    oJobToProc.ImportStatus.Message = "Import data correctly";
                    oJobToProc.ImportStatus.Status = "DONE";
                }


                _db.updateJob(oJobToProc.Serialize(), oJobToProc.JOBID);
            }
           // xlApp.Workbooks.Close();
             * */
        }
    }
}
