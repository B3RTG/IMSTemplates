using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IMSClasses.Jobs
{
    public class OutputParameters
    {
        public class mailparameters
        {
            String MailTo;
            String MailCC;
            String MailBCC;
            String MailFrom;
            String MailBody;
            String MailSubject;
            String MailServer;
            public mailparameters(){}
        }

        public IMSClasses.Jobs.File OriginalFile;
        public String channel;
        public IMSClasses.Jobs.File DestinationFile;
        public mailparameters MailParameters;


        public OutputParameters()
        {
        }

        public bool SetupPath(String sPath)
        {
            this.DestinationFile.Directory = sPath;
            if(this.OriginalFile != null)
                this.OriginalFile.Directory = sPath;


            return true;
        }
        
    }
}
