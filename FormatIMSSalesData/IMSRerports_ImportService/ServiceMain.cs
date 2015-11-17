using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace IMSRerports_ImportService
{
    static class ServiceMain
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main(string[] args)
        {

            if (Environment.UserInteractive)
            {
                ImportService service1 = new ImportService(args);
                service1.TestStartupAndStop(args);
            }
            else
            {
                // Put the body of your old Main method here.
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] 
                { 
                    new ImportService(args) 
                };
                ServiceBase.Run(ServicesToRun);
            }
           
        }
    }
}
