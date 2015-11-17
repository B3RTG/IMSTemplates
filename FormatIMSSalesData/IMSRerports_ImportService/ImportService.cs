using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace IMSRerports_ImportService
{
    public enum ServiceState
    {
        SERVICE_STOPPED = 0x00000001,
        SERVICE_START_PENDING = 0x00000002,
        SERVICE_STOP_PENDING = 0x00000003,
        SERVICE_RUNNING = 0x00000004,
        SERVICE_CONTINUE_PENDING = 0x00000005,
        SERVICE_PAUSE_PENDING = 0x00000006,
        SERVICE_PAUSED = 0x00000007,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ServiceStatus
    {
        public long dwServiceType;
        public ServiceState dwCurrentState;
        public long dwControlsAccepted;
        public long dwWin32ExitCode;
        public long dwServiceSpecificExitCode;
        public long dwCheckPoint;
        public long dwWaitHint;
    };

    public partial class ImportService : ServiceBase
    {
        //Snippet section 16 of code snippet {"project_id":"3fedad16-eaf1-41a6-8f96-0c1949c68f32","entity_id":"db95b2e3-6d54-4438-8e46-53f1cb534551","entity_type":"CodeSnippet","locale":"en-US"} in source file ({"filename":"/CS/MyNewService.cs","blob_type":"Source","blob_id":"-002fcs-002fmynewservice-002ecs","blob_revision":3}) overlaps with other snippet sections. Ensure the tags are placed correctly.


        public ImportService(string[] args)
        {
            InitializeComponent();

            string eventSourceName = "ImportLog_Source";
            string logName = "ImportLog_Source_LOG";
            if (args.Count() > 0) { 
                eventSourceName = args[0]; 
            } if (args.Count() > 1) { 
                logName = args[1]; 
            } 

            ImportServices_EventLog = new EventLog();
            if (!System.Diagnostics.EventLog.SourceExists(eventSourceName))
            {
                System.Diagnostics.EventLog.CreateEventSource(eventSourceName, logName);
            }
            ImportServices_EventLog.Source = eventSourceName;
            ImportServices_EventLog.Log = logName;
        }

        protected override void OnStart(string[] args)
        {
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            ImportServices_EventLog.WriteEntry("In OnStart");
            
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000; // 60 seconds
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();

            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

        }

        protected override void OnStop()
        {
            ImportServices_EventLog.WriteEntry("In OnStop");
        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            ImportServices_EventLog.WriteEntry("Monitoring the system", EventLogEntryType.Information);
        }

        internal void TestStartupAndStop(string[] args)
        {
            this.OnStart(args);
            Console.ReadLine();
            this.OnStop();
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);

    }
}
