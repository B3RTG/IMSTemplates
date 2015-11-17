using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace IMSReportServices
{
    public partial class TaskManager : System.Web.UI.Page
    {
        public const string _FRAME_RESPONSE_VALUE_ = "<textarea data-type=\"application/json\">$RESPONSE$</textarea>";
        private IMSClasses.DBHelper.db oDB;
        private Models.ConfigurationModel oConfig;
        protected void Page_Load(object sender, EventArgs e)
        {
            //Load configuration
            
            oConfig = new Models.ConfigurationModel();
            oDB = new IMSClasses.DBHelper.db(oConfig.ConnectionString);
            
            KeyValuePair<bool, string> oChecks=isCorrectRequest();
            String sResponse = "";

            if(oChecks.Key)
            {
                String sRequestType = Request["requestType"].ToString();
                switch(sRequestType.ToUpper())
                {
                    case "SETTASK":
                        sResponse = CreateTask();
                        Response.ContentType = "text/html";

                        Response.Write(sResponse);
                        Response.End();
                        break;
                    case "GETFILE":
                        GetFile();
                        break;
                }
                
            } 
            else
            {
                sResponse = "{'Status':'ERROR', 'message':'Error en la peticion'}";
                Response.Write(sResponse);
                Response.End();
            }

            
            
            
        }

        private void GetFile()
        {
            Int64 iTaskID = Int64.Parse(Request["TaskID"].ToString());
            String sResponse;
            IMSClasses.Jobs.Job oCurrentJob=null;
            IMSClasses.Jobs.Task oTask=null;

            try
            {
                oTask = oDB.getTaskObject(iTaskID);
                oCurrentJob = oTask.oJob;
            }
            catch (Exception eJobDB)
            {
                //error obteniendo el job de la db
                sResponse = "{\"Status\":\"ERROR\", \"message\":\"Error Obteniendo Task\"}";
            }

            if( oTask!=null && oCurrentJob!=null )
            { //find file
                String sNewPath = System.IO.Path.Combine(oConfig.TaskManager_ProccesFolder , oTask.oJob.JOBCODE.ToString());
                sNewPath = System.IO.Path.Combine(sNewPath, oConfig.TaskManager_ProccesFolder_Out);
                sNewPath = System.IO.Path.Combine(sNewPath, oTask.TaskID.ToString());
                
                if (System.IO.Directory.Exists(sNewPath))
                {
                    String[] sFiles = System.IO.Directory.GetFiles(sNewPath);
                    if(sFiles.Length>0)
                    {//serve first file
                        System.IO.FileInfo oFileInfo = new System.IO.FileInfo(sFiles[0]);
                        if (oFileInfo.Extension.Equals("xls"))
                            Response.ContentType = "application/vnd.ms-excel";
                        else Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                        Response.AppendHeader("content-disposition", "attachment; filename=" + oFileInfo.Name);
                        Response.Clear();
                        Response.WriteFile(sFiles[0]);
                        Response.Flush();
                        Response.End();
                    }
                }
            }
            else
            {
                sResponse = "{\"Status\":\"ERROR\", \"message\":\"Error Obteniendo Task\"}";
            }
        }

        private string CreateTask()
        {
            String sResponse = "";
            //if (sRequestType.ToUpper().Equals("SETTASK") && Request["jobID"] != null)
            //{
            //tenemos que crear una tarea, esto consta de cargar el job pertinente
            //insertar la task de ese job y copiar los ficheros donde toca.
            //--1 obtener JOB para el que se programa la tarea
            Int64 iJobID = Int64.Parse(Request["jobID"].ToString());
            String sJobCode = Request["jobCODE"].ToString();
            IMSClasses.Jobs.Job oCurrentJob = null;
            try
            {
                System.Data.DataRow oJobRow = oDB.getJob(iJobID);
                oCurrentJob = IMSClasses.Jobs.Job.getInstance(oJobRow["JSON"].ToString());
            }
            catch (Exception eJobDB)
            {
                //error obteniendo el job de la db
                sResponse = "{\"Status\":\"ERROR\", \"message\":\"Error Obteniendo Job\"}";
            }


            if (oCurrentJob != null)
            {
                try
                {
                    for (int iFile = 0; iFile < Request.Files.Count; iFile++)
                    {
                        HttpPostedFile oFile = Request.Files[iFile];
                        String sPath = oConfig.UploadPath.Replace("%JOBCODE%", oCurrentJob.JOBCODE);
                        //sPath = System.IO.Path.Combine(sPath, oCurrentJob.JOBCODE);
                        sPath = System.IO.Path.Combine(sPath, oCurrentJob.InputParameters.Files[iFile].Name);//oFile.FileName);
                        if (System.IO.File.Exists(sPath)) System.IO.File.Delete(sPath);
                        oFile.SaveAs(sPath);

                        //oCurrentJob.InputParameters.Files[iFile].Name = oFile.FileName;
                        oCurrentJob.InputParameters.Files[iFile].UploadName = oFile.FileName;
                    }

                    IMSClasses.Jobs.Task oTask = new IMSClasses.Jobs.Task(0, System.DateTime.Now, null, "TODO", "", oCurrentJob.Serialize());
                    IMSClasses.Jobs.Task oNewTask = oDB.CreateTask(oTask);

                    //si se ha creado bien, encolar.
                    if (oNewTask.TaskID > 0)
                    {
                        IMSClasses.RabbitMQ.MessageQueue oQueue = new IMSClasses.RabbitMQ.MessageQueue(oConfig.RabbittMQ.Server, null, oConfig.RabbittMQ.ImportQueueName);
                        oQueue.addMessage((int)oNewTask.TaskID);
                        oQueue.close();

                        sResponse = "{\"Status\":\"OK\", \"message\":\"Tarea creada correctamente.\", \"Task\":" + oTask.Serialize() + "}";
                        sResponse = _FRAME_RESPONSE_VALUE_.Replace("$RESPONSE$", sResponse);
                    }
                    else
                    {
                        sResponse = "{\"Status\":\"ERROR\", \"message\":\"Error creando tareas.\", \"Task\":" + oTask.Serialize() + "}";
                        sResponse = _FRAME_RESPONSE_VALUE_.Replace("$RESPONSE$", sResponse);
                    }
                }
                catch (Exception eCreateTask)
                {
                    sResponse = "{\"Status':\"ERROR\", \"message\":\"Error creando Task:" + eCreateTask.Message + "\"}";
                }

            }
            else
            { //error
                sResponse = "{\"Status':\"ERROR\", \"message\":\"Error obteniendo Job Template\"}";
            }

            //}

            return sResponse;
        }

        private KeyValuePair<bool, string> isCorrectRequest()
        {
            KeyValuePair<bool, string> oResponse = new KeyValuePair<bool, string>(true,"");
            

            if (Request.Params["requestType"] == null || Request.Params["requestType"] == String.Empty)
            {
                oResponse = new KeyValuePair<bool, string>(false, "Request type required.");
            } 
            else
            {
                if (Request.Params["requestType"].ToString().ToUpper().Equals("SETTASK") || Request.Params["requestType"].ToString().ToUpper().Equals("GETFILE"))
                {
                    //correcto
                }
                else
                {
                    oResponse = new KeyValuePair<bool, string>(false, "Action not found.");
                }

                if (oResponse.Value == "" && Request.Params["requestType"].ToString().ToUpper().Equals("SETTASK"))
                {
                    if (Request.Files.Count == 0) oResponse = new KeyValuePair<bool, string>(false, "No files send.");
                    if (Request["jobID"] == null || Request["jobID"].ToString().Equals(String.Empty)) oResponse = new KeyValuePair<bool, string>(false, "Job Id is mandatory.");
                }

                if (oResponse.Value == "" && Request.Params["requestType"].ToString().ToUpper().Equals("GETFILE"))
                {
                    if (Request["TaskID"] == null || Request["TaskID"].ToString().Equals(String.Empty)) oResponse = new KeyValuePair<bool, string>(false, "Task Id is mandatory.");
                }

            }

            

            return oResponse;
        }
    }
}