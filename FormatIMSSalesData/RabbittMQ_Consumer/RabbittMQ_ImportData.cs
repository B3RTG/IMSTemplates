using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using IMSClasses;
using IMSClasses.RabbitMQ;



using System.Reflection;

namespace RabbittMQ_ImportData
{
    class RabbittMQ_ImportData
    {
        public const String _DEFAULT_LOADER_NAME_ = "DefaultLoader";

        static void Main(string[] args)
        {            
            ConfigurationClass oCfg = new ConfigurationClass();
            ICollection<InterfaceDataLoader.IDataLoader> LoadedPlugins;
            Dictionary<string, InterfaceDataLoader.IDataLoader> oPluginList = new Dictionary<string,InterfaceDataLoader.IDataLoader>();

            IMSClasses.RabbitMQ.MessageQueue oRBImportQueue = new MessageQueue(oCfg.RabbittMQ.Server, "", oCfg.RabbittMQ.ImportQueueName);

            LoadedPlugins = RabbittMQ_ImportData.LoadInitialPlugins(oCfg.LoadersDLLDirectory);
            foreach(var oItem in LoadedPlugins)
            {
                if(!oPluginList.ContainsKey(oItem.Name))
                    oPluginList.Add(oItem.Name, oItem);
            }

            while (true)
            {
                Console.WriteLine("Waiting for Jobs.............");
                Int64 iTask = oRBImportQueue.waitConsume();
                bool bCorrect = true;
                String sError = "";

                Console.WriteLine("Executiong of import Task ID --> " + iTask.ToString());

                IMSClasses.DBHelper.db oDB = new IMSClasses.DBHelper.db(oCfg.ConnectionString);
                System.Data.DataRow oRowJob = oDB.getTask(iTask);

                //IMSClasses.Jobs.Job oCurrentJob = IMSClasses.Jobs.Job.getInstance(oRowJob["JSON"].ToString());
                IMSClasses.Jobs.Task oCurrentTask = IMSClasses.Jobs.Task.getInstance(oRowJob["JSON"].ToString());
                if (oCurrentTask.TaskID == 0)
                {
                    oCurrentTask.TaskID = iTask;
                    oDB.updateTask(oCurrentTask);
                }
                if (oCurrentTask.oJob.JOBCODE == null) oCurrentTask.oJob = IMSClasses.Jobs.Job.getInstance(oRowJob["JobJSON"].ToString());
                

                //String sInPath = System.IO.Path.Combine(oCfg.Paths.MainFolder, oCurrentJob.JOBCODE);
                String sInPath = System.IO.Path.Combine(oCfg.Paths.MainFolder, oCurrentTask.oJob.JOBCODE);
                sInPath = System.IO.Path.Combine(sInPath, oCfg.Paths.InFolder);
                //oCurrentJob.InputParameters.SetupInPaths(sInPath);
                oCurrentTask.oJob.InputParameters.SetupInPaths(sInPath);

                //setup folder routes

                InterfaceDataLoader.IDataLoader oCurrentPlugin = null;
                if (oCurrentTask.oJob.PluginName == null ||
                    oCurrentTask.oJob.PluginName.Equals(String.Empty) ||
                    oCurrentTask.oJob.PluginName.Equals(_DEFAULT_LOADER_NAME_)
                  )
                    /*if (oCurrentJob.PluginName == null ||
                        oCurrentJob.PluginName.Equals(String.Empty) ||
                        oCurrentJob.PluginName.Equals(_DEFAULT_LOADER_NAME_)
                      )*/
                { //default loader
                    oCurrentPlugin = oPluginList[_DEFAULT_LOADER_NAME_];
                }
                else
                {
                    if (oPluginList.ContainsKey(oCurrentTask.oJob.PluginName))
                    //if (oPluginList.ContainsKey(oCurrentJob.PluginName))
                    { //lo tenemos, asignar
                        oCurrentPlugin = oPluginList[oCurrentTask.oJob.PluginName];
                        //oCurrentPlugin = oPluginList[oCurrentJob.PluginName];
                    }
                    else
                    {   //intentar cargar nuevos plugins
                        InterfaceDataLoader.IDataLoader LoadedPluging = RabbittMQ_ImportData.LoadPluginByName(_DEFAULT_LOADER_NAME_, oCfg.LoadersDLLDirectory);
                        if(LoadedPluging!= null)
                        {
                            oPluginList.Add(LoadedPluging.Name, LoadedPluging);
                            oCurrentPlugin = LoadedPluging;
                        } else
                        {   //no se ha podido cargar el plugin a utilizar
                            oCurrentPlugin = null;
                            oCurrentTask.oJob.ImportStatus.ExecutionDate = DateTime.Now;
                            oCurrentTask.oJob.ImportStatus.Message = "Plugin for load not found.";
                            oCurrentTask.oJob.ImportStatus.Status = "ERRO";
                            oCurrentTask.StatusCurrent = "ERRO";
                            oCurrentTask.StatusFinal = "ERRO";
                            oCurrentTask.UpdateDate = DateTime.Now;
                            oCurrentTask.TaskComments = "Plugin for load not found.";

                            /*oCurrentJob.ImportStatus.ExecutionDate = DateTime.Now;
                            oCurrentJob.ImportStatus.Message = "Plugin for load not found.";
                            oCurrentJob.ImportStatus.Status = "ERRO";*/
                        }
                    }
                }

                try
                {
                    if (oCurrentPlugin != null){
                        if (oCurrentPlugin.ImportData(oCurrentTask.oJob, oCfg.ConnectionString))
                        {   //done correctly
                            oCurrentTask.oJob.ImportStatus.ExecutionDate = DateTime.Now;
                            oCurrentTask.oJob.ImportStatus.Message = "Import data correctly";
                            oCurrentTask.oJob.ImportStatus.Status = "DONE";
                            oCurrentTask.StatusCurrent = "IMDO";
                            oCurrentTask.UpdateDate = DateTime.Now;
                            oCurrentTask.TaskComments = "Importación de datos correcta";

                            //añadir para formateado
                            IMSClasses.RabbitMQ.MessageQueue oRBExcelQueue = new MessageQueue(oCfg.RabbittMQ.Server, "", oCfg.RabbittMQ.ExcelQueueName);
                            //oRBExcelQueue.addMessage((int)oCurrentJob.JOBID);
                            oRBExcelQueue.addMessage((int)oCurrentTask.TaskID);
                            oRBExcelQueue.close();

                            Console.WriteLine(" <<DONE>> " + oCurrentTask.oJob.ImportStatus.Message + " -- Date: " + oCurrentTask.oJob.ImportStatus.ExecutionDate.ToString());

                        }
                        else
                        {   //error loading
                            oCurrentTask.oJob.ImportStatus.ExecutionDate = DateTime.Now;
                            oCurrentTask.oJob.ImportStatus.Message = sError;
                            oCurrentTask.oJob.ImportStatus.Status = "ERRO";
                            oCurrentTask.StatusCurrent = "ERRO";
                            oCurrentTask.UpdateDate = DateTime.Now;
                            oCurrentTask.TaskComments = sError;
                            Console.WriteLine(" <<Error>> " + sError);
                        }
                    }
                    
                }
                catch (Exception ePluginException)
                {
                    oCurrentTask.oJob.ImportStatus.ExecutionDate = DateTime.Now;
                    oCurrentTask.oJob.ImportStatus.Message = "<Error> Exception Message: " + ePluginException.Message;
                    oCurrentTask.oJob.ImportStatus.Status = "ERRO";
                    oCurrentTask.StatusCurrent = "ERRO";
                    oCurrentTask.UpdateDate = DateTime.Now;
                    oCurrentTask.TaskComments = "<Error> Exception Message: " + ePluginException.Message;
                    Console.WriteLine(" <<Error>> " + sError);
                } 
                finally
                {
                    oRBImportQueue.markLastMessageAsProcessed();
                    if (oCurrentTask != null)
                    {
                        oDB.updateJob(oCurrentTask.oJob.Serialize(), oCurrentTask.oJob.JOBID);
                        oDB.updateTask(oCurrentTask);
                    }
                        
                }                

            }

            /*IMSClasses.RabbitMQ.MessageQueue oRBExcelQueuePurge = new MessageQueue(oCfg.RabbittMQ.Server, "", oCfg.RabbittMQ.ExcelQueueName);
            oRBExcelQueuePurge.purgeQueue();
            oRBExcelQueuePurge.close();*/


            //oRBQueue.purgeQueue();
            //oRBQueue.addMessage(2);

            /*
            while(true)
            {
                Int64 iJobID = oRBImportQueue.waitConsume();
                bool bCorrect = true;
                String sError = "";

                IMSClasses.DBHelper.db oDB = new IMSClasses.DBHelper.db(oCfg.ConnectionString);
                System.Data.DataRow oRowJob = oDB.getJob(iJobID);
                IMSClasses.Jobs.Job oCurrentJob = IMSClasses.Jobs.Job.getInstance(oRowJob["JSON"].ToString());
                System.Data.DataTable dt = null;

                Console.WriteLine("Executiong of import job ID --> " + oCurrentJob.JOBID.ToString());

                try
                {
                    dt = IMSClasses.excellHelpper.ExcelHelpper.getExcelData(oCurrentJob.InputParameters.Files[0].FileName, oCurrentJob.SQLParameters.TableName);
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
                        bCorrect = oDB.LoadTable(dt);
                    }
                    catch (Exception dbException)
                    {
                        bCorrect = false;
                        sError = "Error loading data in DB --> Exception --> " + dbException.Message;
                    }

                }


                oRBImportQueue.markLastMessageAsProcessed();

                if (!bCorrect)
                { //failure update
                    oCurrentJob.ImportStatus.ExecutionDate = DateTime.Now;
                    oCurrentJob.ImportStatus.Message = sError;
                    oCurrentJob.ImportStatus.Status = "ERRO";
                    Console.WriteLine(" <<Error>> " + sError);
                }
                else
                { //correct job update
                    oCurrentJob.ImportStatus.ExecutionDate = DateTime.Now;
                    oCurrentJob.ImportStatus.Message = "Import data correctly";
                    oCurrentJob.ImportStatus.Status = "DONE";

                    //añadir para formateado
                    IMSClasses.RabbitMQ.MessageQueue oRBExcelQueue = new MessageQueue(oCfg.RabbittMQ.Server, "", oCfg.RabbittMQ.ExcelQueueName);
                    oRBExcelQueue.addMessage((int)oCurrentJob.JOBID);
                    oRBExcelQueue.close();

                    Console.WriteLine(" <<DONE>> " + oCurrentJob.ImportStatus.Message + " -- Date: " + oCurrentJob.ImportStatus.ExecutionDate.ToString() );

                }

                

                oDB.updateJob(oCurrentJob.Serialize(), oCurrentJob.JOBID);
                 
            }
            
            */




            //oRBQueue.addMessage(1);
            //oRBQueue.close();

        }

        public static InterfaceDataLoader.IDataLoader LoadPluginByName(String PluginName)
        {
            InterfaceDataLoader.IDataLoader oLoader = null;
            Assembly oAssemblie = null;
            try
            {
                String sDllDirectory = @"C:\Dev\IMS\loaders\";
                String[] dllFileNames = System.IO.Directory.GetFiles(sDllDirectory, "*.dll");

                foreach (string dllFile in dllFileNames)
                {
                    AssemblyName oAsName = AssemblyName.GetAssemblyName(dllFile);
                    if (oAsName.Name.Equals(PluginName))
                    {
                        oAssemblie = Assembly.Load(oAsName);
                    }
                }

                ICollection<Type> pluginTypes = new List<Type>();
                if (oAssemblie != null)
                { //encontrado plugin con ese nombre, intentamos cargar con el interface
                    Type pluginType = typeof(InterfaceDataLoader.IDataLoader);
                    Type[] LoadedAssemblyTypes = oAssemblie.GetTypes();
                    foreach (Type oType in LoadedAssemblyTypes)
                    {
                        if (oType.IsInterface || oType.IsAbstract)
                        {
                            continue;
                        }
                        else
                        {
                            if (oType.GetInterface(pluginType.FullName) != null)
                            {
                                pluginTypes.Add(oType);
                            }
                        }
                    }
                }

                if (pluginTypes.Count > 0)
                { //se a encontrado plugin con ese nombre y el interface de carga definido, cargamos tipo
                    oLoader = (InterfaceDataLoader.IDataLoader)Activator.CreateInstance(pluginTypes.First());
                }
            }
            catch(Exception eLoad) 
            {
                oLoader = null;
            }
            
            return oLoader;
        }

        public static InterfaceDataLoader.IDataLoader LoadPluginByName(String PluginName, String DLLPath)
        {
            InterfaceDataLoader.IDataLoader oLoader = null;
            Assembly oAssemblie = null;
            try
            {
                String sDllDirectory = DLLPath;
                String[] dllFileNames = System.IO.Directory.GetFiles(sDllDirectory, "*.dll");

                foreach (string dllFile in dllFileNames)
                {
                    AssemblyName oAsName = AssemblyName.GetAssemblyName(dllFile);
                    if (oAsName.Name.Equals(PluginName))
                    {
                        oAssemblie = Assembly.Load(oAsName);
                    }
                }

                ICollection<Type> pluginTypes = new List<Type>();
                if (oAssemblie != null)
                { //encontrado plugin con ese nombre, intentamos cargar con el interface
                    Type pluginType = typeof(InterfaceDataLoader.IDataLoader);
                    Type[] LoadedAssemblyTypes = oAssemblie.GetTypes();
                    foreach (Type oType in LoadedAssemblyTypes)
                    {
                        if (oType.IsInterface || oType.IsAbstract)
                        {
                            continue;
                        }
                        else
                        {
                            if (oType.GetInterface(pluginType.FullName) != null)
                            {
                                pluginTypes.Add(oType);
                            }
                        }
                    }
                }

                if (pluginTypes.Count > 0)
                { //se a encontrado plugin con ese nombre y el interface de carga definido, cargamos tipo
                    oLoader = (InterfaceDataLoader.IDataLoader)Activator.CreateInstance(pluginTypes.First());
                }
            }
            catch (Exception eLoad)
            {
                oLoader = null;
            }

            return oLoader;
        }

        public static ICollection<InterfaceDataLoader.IDataLoader> LoadInitialPlugins()
        {
            String sDllDirectory = @"C:\Dev\IMS\loaders\";
            String[] dllFileNames = System.IO.Directory.GetFiles(sDllDirectory, "*.dll");

            ICollection<Assembly> assemblies = new List<Assembly>(sDllDirectory.Length);
            foreach(string dllFile in dllFileNames)
            {
                AssemblyName oAsName = AssemblyName.GetAssemblyName(dllFile);
                Assembly oAssembly = Assembly.Load(oAsName);
                assemblies.Add(oAssembly);
            }

            Type pluginType = typeof(InterfaceDataLoader.IDataLoader);
            ICollection<Type> pluginTypes = new List<Type>(); 
            foreach(Assembly oLoadedAssembly in assemblies)
            {
                Type[] LoadedAssemblyTypes = oLoadedAssembly.GetTypes();
                foreach(Type oType in LoadedAssemblyTypes)
                {
                    if(oType.IsInterface || oType.IsAbstract)
                    {
                        continue;
                    }
                    else
                    {
                        if(oType.GetInterface(pluginType.FullName) != null)
                        {
                            pluginTypes.Add(oType);
                        }
                    }
                }
            }

            //ahora podemos crear instancias a partir de los tipos encontrados.
            ICollection<InterfaceDataLoader.IDataLoader> plugins = new List<InterfaceDataLoader.IDataLoader>(pluginTypes.Count); 
            foreach(Type oType in pluginTypes)
            {
                InterfaceDataLoader.IDataLoader oLoader = (InterfaceDataLoader.IDataLoader) Activator.CreateInstance(oType);
                plugins.Add(oLoader);
            }

            return plugins;
        }

        public static ICollection<InterfaceDataLoader.IDataLoader> LoadInitialPlugins(String DLLPath)
        {
            String sDllDirectory = DLLPath;
            String[] dllFileNames = System.IO.Directory.GetFiles(sDllDirectory, "*.dll");

            ICollection<Assembly> assemblies = new List<Assembly>(sDllDirectory.Length);
            foreach (string dllFile in dllFileNames)
            {
                AssemblyName oAsName = AssemblyName.GetAssemblyName(dllFile);
                Assembly oAssembly = Assembly.Load(oAsName);
                assemblies.Add(oAssembly);
            }

            Type pluginType = typeof(InterfaceDataLoader.IDataLoader);
            ICollection<Type> pluginTypes = new List<Type>();
            foreach (Assembly oLoadedAssembly in assemblies)
            {
                Type[] LoadedAssemblyTypes = oLoadedAssembly.GetTypes();
                foreach (Type oType in LoadedAssemblyTypes)
                {
                    if (oType.IsInterface || oType.IsAbstract)
                    {
                        continue;
                    }
                    else
                    {
                        if (oType.GetInterface(pluginType.FullName) != null)
                        {
                            pluginTypes.Add(oType);
                        }
                    }
                }
            }

            //ahora podemos crear instancias a partir de los tipos encontrados.
            ICollection<InterfaceDataLoader.IDataLoader> plugins = new List<InterfaceDataLoader.IDataLoader>(pluginTypes.Count);
            foreach (Type oType in pluginTypes)
            {
                InterfaceDataLoader.IDataLoader oLoader = (InterfaceDataLoader.IDataLoader)Activator.CreateInstance(oType);
                plugins.Add(oLoader);
            }

            return plugins;
        }
    }
}
