using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using IMSClasses.excellHelpper;
using IMSClasses.DBHelper;


namespace GALoader
{
    public class GALoader : InterfaceDataLoader.IDataLoader
    {
        private const String PluginName = "GALoader";

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
            db oDB = new db(DBConnectionString);
            int iIdentity = 1;

            List<IMSClasses.Jobs.File>.Enumerator oEFiles = oCurrentJob.InputParameters.Files.GetEnumerator();

            while (oEFiles.MoveNext() && bCorrect)
            {
                try
                {
                    String sTableName = oCurrentJob.SQLParameters.TableName.Replace(@"%identity%", iIdentity.ToString());
                    IMSClasses.Jobs.File oCurrentFile = oEFiles.Current;
                    oDT = ExcelHelpper.getExcelData(oCurrentFile.FileName, sTableName);
                } catch(Exception eLoadException)
                {
                    bCorrect = false;
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
                    }
                }
                if (bCorrect) iIdentity++;


                
            }



            return bCorrect;
        }

    }
}
