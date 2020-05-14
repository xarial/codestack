using CodeStack.SwEx.AddIn;
using CodeStack.SwEx.AddIn.Attributes;
using CodeStack.SwEx.Common.Attributes;
using System;
using System.Runtime.InteropServices;

namespace CodeStack.SwEx
{
    [LoggerOptions(true, LOGGER_NAME)]
    [AutoRegister]
    [ComVisible(true), Guid("CD7C743A-3C82-4A4A-B557-BBD6228CE2C8")]
    public class LogAddIn : SwAddInEx
    {
        internal const string LOGGER_NAME = "MyAddInLog";

        public override bool OnConnect()
        {
            try
            {
                Logger.Log("Loading add-in...");

                //TODO: implement connection
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                throw;
            }
        }
    }
}
