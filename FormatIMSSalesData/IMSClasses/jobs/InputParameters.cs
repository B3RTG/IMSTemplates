using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IMSClasses.Jobs
{
    public class InputParameters
    {
        public List<IMSClasses.Jobs.File> Files;
        public IMSClasses.Jobs.File TemplateFile;

        public InputParameters() 
        {
            this.Files = new List<IMSClasses.Jobs.File>();
            this.TemplateFile = new IMSClasses.Jobs.File();
        }

        public bool SetupInPaths(String sPath)
        {

            foreach(Jobs.File oFile in this.Files)
            {
                oFile.Directory = sPath;
            }

            return true;
        }

        public bool SetupTemplatePaths(String sPath)
        {

            TemplateFile.Directory = sPath;
            return true;
        }
    }
}
