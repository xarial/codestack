using SolidWorks.Interop.sldworks;
using System;
using System.Linq;

namespace CodeStack.GeometryHelper
{
    internal class GeometryHelperService
    {
        private readonly ISldWorks m_App;

        internal GeometryHelperService(ISldWorks app)
        {
            m_App = app;
        }

        internal int GetFacesCountFromSelectedBody(double minArea)
        {
            var model = m_App.IActiveDoc2;

            if (model != null)
            {
                var body = model.ISelectionManager.GetSelectedObject6(1, -1) as IBody2;

                if (body != null)
                {
                    var faces = body.GetFaces() as object[];

                    if (faces != null)
                    {
                        return faces.Count(f => (f as IFace2).GetArea() >= minArea);
                    }
                    else
                    {
                        throw new NullReferenceException("No faces in the body");
                    }
                }
                else
                {
                    throw new NullReferenceException("Body is not selected");
                }
            }
            else
            {
                throw new NullReferenceException("Model is not opened");
            }
        }
    }
}
