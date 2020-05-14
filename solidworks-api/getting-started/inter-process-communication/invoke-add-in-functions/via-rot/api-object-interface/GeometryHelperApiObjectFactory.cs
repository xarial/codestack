using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace CodeStack.GeometryHelper
{
    [ComVisible(true)]
    public interface IGeometryHelperApiObjectFactory
    {
        string GetName(int prcId);
        IGeometryHelperApiObject GetInstance(int prcId);
    }

    [ComVisible(true)]
    [ProgId("GeometryHelper.ApiObjectFactory")]
    public class GeometryHelperApiObjectFactory : IGeometryHelperApiObjectFactory
    {
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        public string GetName(int prcId)
        {
            return $"GeometryHelperApiObjectFactory_PID_{prcId}";
        }

        public IGeometryHelperApiObject GetInstance(int prcId)
        {
            return FindObjectByMonikerName<IGeometryHelperApiObject>($"!{GetName(prcId)}");
        }

        private T FindObjectByMonikerName<T>(string monikerName)
            where T : class
        {
            IBindCtx context = null;
            IRunningObjectTable rot = null;
            IEnumMoniker monikers = null;

            try
            {
                CreateBindCtx(0, out context);

                context.GetRunningObjectTable(out rot);

                rot.EnumRunning(out monikers);

                var moniker = new IMoniker[1];

                while (monikers.Next(1, moniker, IntPtr.Zero) == 0)
                {
                    var curMoniker = moniker.First();

                    string name = null;

                    if (curMoniker != null)
                    {
                        try
                        {
                            curMoniker.GetDisplayName(context, null, out name);
                        }
                        catch (UnauthorizedAccessException)
                        {
                        }
                    }

                    if (string.Equals(monikerName,
                        name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        object app;
                        rot.GetObject(curMoniker, out app);
                        return (T)app;
                    }
                }
            }
            finally
            {
                if (monikers != null)
                {
                    Marshal.ReleaseComObject(monikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (context != null)
                {
                    Marshal.ReleaseComObject(context);
                }
            }

            return null;
        }
    }
}
