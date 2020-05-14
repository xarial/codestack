using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace CodeStack.GeometryHelper
{
    public static class RotHelper
    {   
        [DllImport("ole32.dll", ExactSpelling = true, PreserveSig = false)]
        private static extern IRunningObjectTable GetRunningObjectTable(
            int reserved);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode,
             ExactSpelling = true, PreserveSig = false)]
        private static extern IMoniker CreateItemMoniker(
            [In] string lpszDelim, [In] string lpszItem);

        public static void Register(object obj, string name)
        {
            IRunningObjectTable rot = null;
            IMoniker moniker = null;

            try
            {
                rot = GetRunningObjectTable(0);

                moniker = CreateItemMoniker("!", name);

                const int ROTFLAGS_REGISTRATIONKEEPSALIVE = 1;
                var cookie = rot.Register(ROTFLAGS_REGISTRATIONKEEPSALIVE, obj, moniker);
            }
            finally
            {
                if (moniker != null)
                {
                    Marshal.ReleaseComObject(moniker);
                }
                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }
            }
        }
    }
}
