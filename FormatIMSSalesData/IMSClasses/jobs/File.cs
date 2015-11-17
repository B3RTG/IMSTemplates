using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IMSClasses.Jobs
{
    public class File
    {
        public String SampleName { get; set; }
        public String Name { get; set; }
        public String Directory { get; set; }

        public String UploadName { get; set; }
        public String GUIDName { get; set; }

        public File() { }

        public File(String sFileName, String sFilePath)
        {
            this.Name = sFileName;
            this.Directory = sFilePath;
        }

        public String FileName
        {
            get
            {
                String sResult = "";
                if (this.Directory != null && this.Name != null)
                    sResult = System.IO.Path.Combine(this.Directory, this.Name);
                return sResult;
            }
        }
    }
}
