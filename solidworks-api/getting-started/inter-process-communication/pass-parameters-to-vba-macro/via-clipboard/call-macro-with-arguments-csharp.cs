using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;
using System.Windows.Forms;
using System.Text;
using System.IO;

namespace CodeStack
{
    public partial class SolidWorksMacro
    {
        const string ARG_NAME = "__SwMacroArgs__";

        public void Main()
        {
            SetArgument("Argument from C# macro");

            int err;
            if (!swApp.RunMacro2(@"D:\Macros\GetArgumentMacro.swp",
                "Macro1", "main", (int)swRunMacroOption_e.swRunMacroUnloadAfterRun, out err))
            {
                swApp.SendMsgToUser(string.Format("Failed to run macro. Error code: {0}", err));
            }
        }

        private static void SetArgument(string arg)
        {
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(arg)))
            {
                Clipboard.SetData(ARG_NAME, stream);
            }
        }

        public SldWorks swApp;
    }
}


