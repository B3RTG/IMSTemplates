using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.SqlClient;

namespace IMSClasses.LogHelpper
{
    public class LogHelper
    {
        private String sLogPath;
        private String sAppName;

        public LogHelper(String sLogPath, String sAppName)
        {
            this.sLogPath = sLogPath;
            this.sAppName = sAppName;
            TextFile_OpenLog(sLogPath, sAppName);
        }

        public void TextFile_OpenLog(String sLogPath, String sAppName)
        {
            StreamWriter sWriter = new StreamWriter(sLogPath, true);
            sWriter.WriteLine(DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + " <START> " + sAppName);
            sWriter.Close();
        }

        public void TextFile_addLogLine(String sMessage, int Tabs, String sScope)
        {
            String sLine = DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + " ";
            if (Tabs > 0)
            {
                for (int i = 0; i <= Tabs; i++)
                {
                    sLine += "\t";
                }
            }
            sLine += " <" + sScope + "> " + sMessage;
            StreamWriter sWriter = new StreamWriter(sLogPath, true);
            sWriter.WriteLine(sLine);
            sWriter.Close();
        }

        public void TextFile_CloseLog()
        {
            StreamWriter sWriter = new StreamWriter(sLogPath, true);
            sWriter.WriteLine(DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + " </START>");
            sWriter.Close();
        }


        public static bool DB_AddLine(String sConnectionString, String[] Parameters, String[] values, String sProcedure)
        {
            bool bResult = true;
            try
            {
                SqlConnection oConnection = new SqlConnection(sConnectionString);
                SqlCommand oCommand = new SqlCommand(sProcedure, oConnection);
                int iParameter = 0;


                foreach (String sParameter in Parameters)
                {
                    oCommand.Parameters.AddWithValue(sParameter, values[iParameter]);
                }


                oCommand.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                bResult = false;
            }

            return bResult;
        }
    }
}
