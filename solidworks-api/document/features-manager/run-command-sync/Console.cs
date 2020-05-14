using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using System;
using System.Threading.Tasks;

namespace RunCommandAsyncConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            AsyncMain().Wait();
            return;
        }

        static async Task AsyncMain()
        {
            var app = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as ISldWorks;
            app.Visible = true;

            var res = await app.RunCommandAsync(swCommands_e.swCommands_Extrude);

            if (res)
            {
                app.SendMsgToUser("Command Completed");
            }
            else
            {
                app.SendMsgToUser("Command Canceled");
            }
        }
    }
}
