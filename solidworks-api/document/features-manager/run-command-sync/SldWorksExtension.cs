using SolidWorks.Interop.swcommands;
using System.Threading.Tasks;

namespace SolidWorks.Interop.sldworks
{
    public static class SldWorksExtension
    {
        public static Task<bool> RunCommandAsync(this ISldWorks app, swCommands_e cmd)
        {
            return Task.Run(() => 
            {
                if (app.RunCommand((int)cmd, ""))
                {
                    var isCmdCompleted = false;
                    var res = false;

                    (app as SldWorks).CommandCloseNotify += (int Command, int reason) =>
                    {
                        res = reason == (int)swCommands_e.swCommands_PmOK;
                        isCmdCompleted = true;
                        return 0;
                    };

                    while (!isCmdCompleted)
                    {
                        Task.Delay(10);
                    }

                    return res;
                }

                return false;
            });
        }
    }
}
