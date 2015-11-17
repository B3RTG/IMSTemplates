using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace TestTemplate
{
    [ComVisible(true)]
    [Guid("B523844E-1A41-4118-A0F0-FDFA7BCD77C9")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITestClass
    {
        void setupTask();
    }

    public class TestClass : ITestClass
    {
        public void setupTask()
        {
            Globals.ThisWorkbook.TaskID = 1;
            System.Windows.Forms.MessageBox.Show("Test");
        }
    }
}
