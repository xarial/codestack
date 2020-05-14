using System.Runtime.InteropServices;

namespace CodeStack.GeometryHelper
{
    [ComVisible(true)]
    public interface IGeometryHelperApiObject
    {
        int GetFacesCount(double minArea);
    }
}
