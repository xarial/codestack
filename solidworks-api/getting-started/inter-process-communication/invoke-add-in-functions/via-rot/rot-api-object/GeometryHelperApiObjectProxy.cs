using System;
using System.Runtime.InteropServices;

namespace CodeStack.GeometryHelper
{
    [ComVisible(true)]
    public class GeometryHelperApiObjectProxy : IGeometryHelperApiObject
    {
        internal event Func<double, int> GetFacesCountRequested;

        public int GetFacesCount(double minArea)
        {
            if (GetFacesCountRequested != null)
            {
                return GetFacesCountRequested.Invoke(minArea);
            }
            else
            {
                throw new Exception("API object not connected");
            }
        }
    }
}
