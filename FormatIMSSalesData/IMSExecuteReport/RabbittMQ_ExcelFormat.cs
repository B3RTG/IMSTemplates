using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using IMSClasses.RabbitMQ;
using IMSClasses.excellHelpper;

namespace RabbittMQ_ExcelFormat
{
    class RabbittMQ_ExcelFormat
    {
        static void Main(string[] args)
        {
            ConfigurationClass oCfg = new ConfigurationClass();

            IMSClasses.RabbitMQ.MessageQueue oRBFormatQueue = new MessageQueue(oCfg.RabbittMQ.Server, "", oCfg.RabbittMQ.ExcelQueueName);


            /*IMSClasses.RabbitMQ.MessageQueue oRBExcelQueuePurge = new MessageQueue(oCfg.RabbittMQ.Server, "", oCfg.RabbittMQ.ExcelQueueName);
            oRBExcelQueuePurge.purgeQueue();
            oRBExcelQueuePurge.close();*/


            //oRBQueue.purgeQueue();
            //oRBQueue.addMessage(2);
            while (true)
            {
                Console.WriteLine("Waiting for Jobs.............");
                Int64 iTaskID = oRBFormatQueue.waitConsume();
                bool bCorrect = true;
                String sError = "";
                System.Data.DataTable dt = null;
                
                IMSClasses.DBHelper.db oDB = new IMSClasses.DBHelper.db(oCfg.ConnectionString);
                System.Data.DataRow oRowTask = null;
                try
                {
                    oRowTask = oDB.getTask(iTaskID);
                } catch(Exception eNotFound)
                {
                    bCorrect = false;
                }

                if(bCorrect)
                {
                    IMSClasses.Jobs.Task oCurrentTask = IMSClasses.Jobs.Task.getInstance(oRowTask["JSON"].ToString());
                    try
                    {
                        String sTemplatePath = System.IO.Path.Combine(oCfg.Paths.MainFolder, oCurrentTask.oJob.JOBCODE);
                        String sOutputPath = System.IO.Path.Combine(oCfg.Paths.MainFolder, oCurrentTask.oJob.JOBCODE);
                        sTemplatePath = System.IO.Path.Combine(sTemplatePath, oCfg.Paths.TemplateFolder);
                        sOutputPath = System.IO.Path.Combine(sOutputPath, oCfg.Paths.OutFolder);
                        oCurrentTask.oJob.InputParameters.SetupTemplatePaths(sTemplatePath);
                        oCurrentTask.oJob.OutputParameters.SetupPath(sOutputPath);


                        //String sJson = oCurrentJob.Serialize();
                        Console.WriteLine("Executiong of Format for job ID --> " + oCurrentTask.TaskID.ToString());

                        try
                        {
                            //dt = IMSClasses.excellHelpper.ExcelHelpper.getExcelData(oCurrentJob.InputParameters.Files[0].FileName, oCurrentJob.SQLParameters.TableName);
                            //ExcelHelpper.executeExcelTemplate(@"C:\Dev\IMS\bin\Templates\Top50Farmacias\Top50Farmacias\bin\Release\Top50Farmacias.xltx");
                            ExcelHelpper.executeExcelTemplate(oCurrentTask.oJob.InputParameters.TemplateFile.FileName);
                        }
                        catch (Exception xlException)
                        {
                            bCorrect = false;
                            sError = "Error formatting excel --> Exception --> " + xlException.Message;
                        }


                        oRBFormatQueue.markLastMessageAsProcessed();

                        //oRowTask = oDB.getTask(iTaskID);
                        oRowTask = oDB.getJob(oCurrentTask.oJob.JOBID);
                        IMSClasses.Jobs.Job oCurrentJob = IMSClasses.Jobs.Job.getInstance(oRowTask["JSON"].ToString());

                        if (!bCorrect || oCurrentJob.ReportStatus.Status.Equals("ERRO"))
                        { //failure update
                            /*oCurrentJob.ReportStatus.ExecutionDate = DateTime.Now;
                            oCurrentJob.ReportStatus.Message = sError;
                            oCurrentJob.ReportStatus.Status = "ERRO";*/
                            Console.WriteLine(" <<Error>> " + sError);
                            oCurrentTask.UpdateDate = DateTime.Now;
                            oCurrentTask.TaskComments = oCurrentJob.ReportStatus.Message;
                            oCurrentTask.StatusFinal = oCurrentJob.ReportStatus.Status;
                            oCurrentTask.StatusCurrent = oCurrentJob.ReportStatus.Status;
                        }
                        else
                        { //correct job update
                            /*como vamos por task y queremos manterlo estructurado en el servidor, vamos a mover el fichero generado*/
                            foreach (String sFile in System.IO.Directory.GetFiles(oCurrentJob.OutputParameters.DestinationFile.Directory))
                            {
                                System.IO.FileInfo oFileInfo = new System.IO.FileInfo(sFile);
                                String sNewPath = System.IO.Path.Combine(oCurrentJob.OutputParameters.DestinationFile.Directory, oCurrentTask.TaskID.ToString());
                                if (!System.IO.Directory.Exists(sNewPath)) System.IO.Directory.CreateDirectory(sNewPath);
                                oFileInfo.MoveTo(System.IO.Path.Combine(sNewPath, oFileInfo.Name));
                            }


                            oCurrentTask.oJob.ReportStatus.ExecutionDate = DateTime.Now;
                            oCurrentTask.oJob.ReportStatus.Message = "Format excel correctly";
                            oCurrentTask.oJob.ReportStatus.Status = "DONE";

                            oCurrentTask.UpdateDate = DateTime.Now;
                            oCurrentTask.TaskComments = "Format excel correctly";
                            oCurrentTask.StatusFinal = "DONE";
                            oCurrentTask.StatusCurrent = "DONE";

                            Console.WriteLine(" <<DONE>> " + oCurrentTask.StatusCurrent + " -- Date: " + oCurrentTask.UpdateDate.ToString());

                        }

                        oDB.updateJob(oCurrentTask.oJob.Serialize(), oCurrentTask.oJob.JOBID);
                        oDB.updateTask(oCurrentTask);
                    }
                    catch (Exception eTaskProccess)
                    {
                        oCurrentTask.StatusCurrent = "ERRO";
                        oCurrentTask.StatusFinal = "ERRO";
                        oCurrentTask.TaskComments = "ERROR >> " + eTaskProccess.Message.ToString();
                        oDB.updateTask(oCurrentTask);
                    }
                }
                else
                {
                    oRBFormatQueue.markLastMessageAsProcessed();
                }

            }

        }
    }
}
