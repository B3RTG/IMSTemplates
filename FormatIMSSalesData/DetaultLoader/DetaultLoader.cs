using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.Data;
using IMSClasses.excellHelpper;
using IMSClasses.DBHelper;


namespace DefaultLoader
{
    public class DefaultLoader : InterfaceDataLoader.IDataLoader
    {
        private const String PluginName = "DefaultLoader";


        public string Name
        {
            get
            {
                return PluginName;
            }
        }

        public bool ImportData(Object Configuration, String DBConnectionString)
        {

            IMSClasses.Jobs.Job oCurrentJob = (IMSClasses.Jobs.Job)Configuration;
            
            DataTable oDT = null;
            bool bCorrect = true;
            String sError = "";
            db oDB = new db(DBConnectionString);
            try
            {
                oDT = IMSClasses.excellHelpper.ExcelHelpper.getExcelData(oCurrentJob.InputParameters.Files[0].FileName, oCurrentJob.SQLParameters.TableName);
            }
            catch (Exception xlException)
            {
                bCorrect = false;
                sError = "Error getting excel data --> Exception --> " + xlException.Message;
            }

            if (bCorrect)
            {
                try
                {
                    bCorrect = oDB.LoadTable(oDT);
                }
                catch (Exception dbException)
                {
                    bCorrect = false;
                    sError = "Error loading data in DB --> Exception --> " + dbException.Message;
                }

            }
            

            
            return bCorrect;
        }
    }
}
