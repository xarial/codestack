using System.Runtime.InteropServices;

namespace CodeStack.GeometryHelper
{
    [ComVisible(true)]
    public class GeometryHelperApiObject : IGeometryHelperApiObject
    {
        private readonly GeometryHelperService m_GeomSvc;

        internal GeometryHelperApiObject(GeometryHelperService geomSvc)
        {
            m_GeomSvc = geomSvc;
        }

        public int GetFacesCount(double minArea)
        {
            return m_GeomSvc.GetFacesCountFromSelectedBody(minArea);
        }
    }
}
