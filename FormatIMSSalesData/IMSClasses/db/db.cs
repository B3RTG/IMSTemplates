using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace IMSClasses.DBHelper
{
    

    public class db
    {
        public String ConnectionString;
        public SqlConnection oConnection;
        private SqlTransaction Transaction;
        //private String server = @"Data Source=pdwesbar01\sql2005;Initial Catalog=PEDRO;User Id=sa;Password=Marina1618;";

        #region DB_METHODS
        public db(String sConnectionString)
        {
            this.ConnectionString = sConnectionString;
        }

        protected bool ConnectionStart()
        {
            bool bResult = true;
            try
            {
                if (this.oConnection == null) this.oConnection = new SqlConnection(this.ConnectionString);
                if (this.oConnection.State != System.Data.ConnectionState.Open) this.oConnection.Open();
            }
            catch (Exception e)
            {
                bResult = false;
            }
            return bResult;
        }

        protected bool ConectionClose()
        {
            bool bResult = true;
            try
            {
                if (this.oConnection.State == System.Data.ConnectionState.Open) this.oConnection.Close();
            }
            catch (Exception e)
            {
                bResult = false;
            }
            return bResult;
        }

        public bool TransactionStart()
        {
            bool bConnected = false;
            bool bTransanctionOpen = false;
            if (this.oConnection == null)
            {
                bConnected = this.ConnectionStart();
            }
            else if (this.oConnection.State == System.Data.ConnectionState.Open)
            {
                bConnected = true;
            }
            else
            {
                try
                {
                    this.oConnection.Open();
                    bConnected = true;
                }
                catch
                {
                    bConnected = false;
                }

            }

            if (bConnected)
            {
                try
                {
                    this.Transaction = this.oConnection.BeginTransaction();
                    bTransanctionOpen = true;
                }
                catch (Exception e)
                {
                    bTransanctionOpen = false;
                }
            }


            return bTransanctionOpen;
        }
        public bool TransactionCommint()
        {
            if (this.Transaction != null)
            {
                this.Transaction.Commit();
                this.ConectionClose();
            }
            return true;
        }

        public bool TransactionRollBack()
        {
            if (this.Transaction != null)
            {
                this.Transaction.Rollback();
                this.ConectionClose();
            }
            return true;
        }

        public DataSet ExecuteQuery(String sSQLQuery, CommandType oCommandType, List<Object> Parameters = null, String ParameterSufix = "p")
        {
            DataSet oDataSet = new DataSet();

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand(sSQLQuery, this.oConnection);
                oCmd.CommandType = oCommandType;

                if (Parameters != null)
                {
                    int i = 0;
                    foreach (var oItem in Parameters)
                    {
                        String sParamName = "@" + ParameterSufix + i.ToString();
                        oCmd.Parameters.AddWithValue(sParamName, oItem.ToString());

                        i++;
                    }
                }


                SqlDataAdapter oDataAdapter = new SqlDataAdapter(oCmd);
                oDataAdapter.Fill(oDataSet);


                this.ConectionClose();

            }

            return oDataSet;
        }

        #endregion DB_METHODS

        #region JOBS
        public DataTable getConfiguration()
        {
            DataSet oDataSet = new DataSet();

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Get_Configuration", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter oDataAdapter = new SqlDataAdapter(oCmd);
                
                oDataAdapter.Fill(oDataSet);


                this.ConectionClose();

            }

            return oDataSet.Tables[0];
        }

        public bool updateJob(String JSON, Int64 JobID)
        {

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Job_Update", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;

                oCmd.Parameters.AddWithValue("@Json", JSON);
                oCmd.Parameters.AddWithValue("@JobID", JobID);

                oCmd.ExecuteNonQuery();

                this.ConectionClose();

            }

            return true;
        }

        public bool updateJob(IMSClasses.Jobs.Job oJob)
        {

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Job_Update", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;

                oCmd.Parameters.AddWithValue("@Json", oJob.Serialize());
                oCmd.Parameters.AddWithValue("@JobID", oJob.JOBID);

                oCmd.ExecuteNonQuery();

                this.ConectionClose();

            }

            return true;
        }

        

        public DataRow getJob(Int64 JobID)
        {
            DataSet oDataSet = new DataSet();

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Job_get", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;

                oCmd.Parameters.AddWithValue("@JobID", JobID);

                SqlDataAdapter oDataAdapter = new SqlDataAdapter(oCmd);

                oDataAdapter.Fill(oDataSet);


                this.ConectionClose();

            }

            return oDataSet.Tables[0].Rows[0];
        }
        #endregion

        #region TASK
        public DataTable getPendingTask()
        {
            DataSet oDataSet = new DataSet();

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Task_GetPending", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter oDataAdapter = new SqlDataAdapter(oCmd);

                oDataAdapter.Fill(oDataSet);


                this.ConectionClose();

            }

            return oDataSet.Tables[0];
        }

        public DataTable getTaskList()
        {
            DataSet oDataSet = new DataSet();

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Task_GetList", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter oDataAdapter = new SqlDataAdapter(oCmd);

                oDataAdapter.Fill(oDataSet);


                this.ConectionClose();

            }

            return oDataSet.Tables[0];
        }

        public DataRow getTask(Int64 iTaskID)
        {
            DataSet oDataSet = new DataSet();

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Task_get", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;

                oCmd.Parameters.AddWithValue("@TaskID", iTaskID);

                SqlDataAdapter oDataAdapter = new SqlDataAdapter(oCmd);

                oDataAdapter.Fill(oDataSet);


                this.ConectionClose();

            }

            return oDataSet.Tables[0].Rows[0];
        }

        public IMSClasses.Jobs.Task getTaskObject(Int64 iTaskID)
        {
            DataSet oDataSet = new DataSet();

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Task_get", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;

                oCmd.Parameters.AddWithValue("@TaskID", iTaskID);

                SqlDataAdapter oDataAdapter = new SqlDataAdapter(oCmd);

                oDataAdapter.Fill(oDataSet);


                this.ConectionClose();

            }

            IMSClasses.Jobs.Task oTask = null;
            if (oDataSet.Tables[0].Rows.Count > 0)
                oTask = IMSClasses.Jobs.Task.getInstance(oDataSet.Tables[0].Rows[0]["JSON"].ToString());
            return oTask;
        }

        public bool updateTask(IMSClasses.Jobs.Task oTask)
        {

            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Task_Update", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;
                String FinalStatus = "", Comments = "";
                if (oTask.StatusFinal != null) FinalStatus = oTask.StatusFinal;
                if (oTask.TaskComments != null) Comments = oTask.TaskComments;

                oCmd.Parameters.AddWithValue("@Json", oTask.Serialize());
                oCmd.Parameters.AddWithValue("@TaskID", oTask.TaskID);
                oCmd.Parameters.AddWithValue("@EndStatus", FinalStatus);
                oCmd.Parameters.AddWithValue("@CurrentStatus", oTask.StatusCurrent);
                oCmd.Parameters.AddWithValue("@Message", Comments);



                oCmd.ExecuteNonQuery();

                this.ConectionClose();

            }

            return true;
        }

        public IMSClasses.Jobs.Task CreateTask(IMSClasses.Jobs.Task oNewTask)
        {
            IMSClasses.Jobs.Task oResponse = oNewTask;
            if (this.ConnectionStart())
            {
                SqlCommand oCmd = new SqlCommand("Task_Create", this.oConnection);
                oCmd.CommandType = CommandType.StoredProcedure;
                String FinalStatus = "", Comments = "";
                if (oNewTask.StatusFinal != null) FinalStatus = oNewTask.StatusFinal;
                if (oNewTask.TaskComments != null) Comments = oNewTask.TaskComments;

                oCmd.Parameters.AddWithValue("@Json", oNewTask.Serialize());
                oCmd.Parameters.AddWithValue("@JobID", oNewTask.oJob.JOBID);
                oCmd.Parameters.AddWithValue("@CreateDate", oNewTask.CreateDate);
                oCmd.Parameters.AddWithValue("@CurrentStatus", oNewTask.StatusCurrent);
                oCmd.Parameters.AddWithValue("@Message", Comments);
                SqlParameter pTaskID = oCmd.Parameters.Add("@TaskID", SqlDbType.BigInt);
                pTaskID.Direction = ParameterDirection.Output;
                try
                {
                    oCmd.ExecuteNonQuery();
                    oResponse.TaskID = Int64.Parse(oCmd.Parameters["@TaskID"].Value.ToString());
                }
                catch
                {
                    //error creando tarea
                    oResponse.TaskID = -1;
                }
                

                this.ConectionClose();
            }

            return oResponse;
        }

        #endregion

        #region SQLLOAD
        public bool LoadTable(DataTable dt)
        {
            String exists = null;
            try
            {

                if (this.ConnectionStart())
                {
                    StringBuilder oSBTable = new StringBuilder();
                    //1 add drop if exist.
                    oSBTable.Append("IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[");
                    oSBTable.Append(dt.TableName);
                    oSBTable.Append("]') AND type in (N'U')) DROP TABLE [dbo].[");
                    oSBTable.Append(dt.TableName);
                    oSBTable.Append("]; ");
                    oSBTable.AppendLine("");
                    oSBTable.Append("CREATE TABLE [");
                    oSBTable.Append(dt.TableName);
                    //oSBTable.Append("] ( PK_Order bigint identity(1,1),");
                    oSBTable.Append("] ( ");

                    Boolean bFirst = true;
                    foreach (DataColumn dc in dt.Columns)
                    {
                        oSBTable.Append("[");
                        oSBTable.Append(dc.ColumnName);
                        if (bFirst)
                        {
                            oSBTable.Append("]  varchar(MAX),");
                            bFirst = false;
                        }
                        else oSBTable.Append("]  numeric(30,10),");


                    }
                    //oSBTable.Remove(oSBTable.Length-1, 1);
                    oSBTable.Append("PK_Order bigint identity(1,1)");
                    oSBTable.Append(")");
                    SqlCommand cmdCreateTable = new SqlCommand(oSBTable.ToString(), this.oConnection);
                    cmdCreateTable.ExecuteNonQuery();

                    // copying the data from datatable to database table
                    using (SqlBulkCopy bulkcopy = new SqlBulkCopy(this.oConnection))
                    {
                        bulkcopy.DestinationTableName = "[" + dt.TableName + "]";
                        bulkcopy.WriteToServer(dt);
                    }
                }
                else
                {
                    throw new Exception("Problems with DB Connection");
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return true;
        }
        #endregion
        


    }
}
