using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;

namespace HoleOrBoss.csproj
{
    public partial class SolidWorksMacro
    {
        public void Main()
        {
            IModelDoc2 doc = swApp.IActiveDoc2;

            if (doc != null)
            {
                IFace2 face = doc.ISelectionManager.GetSelectedObject6(1, -1) as IFace2;

                if (face != null)
                {
                    if (IsHole(face))
                    {
                        swApp.SendMsgToUser("Selected face is hole");
                    }
                    else
                    {
                        swApp.SendMsgToUser("Selected face is boss");
                    }
                }
                else
                {
                    throw new Exception("Face is not selected");
                }
            }
            else
            {
                throw new Exception("No document opened");
            }
        }

        private bool IsHole(IFace2 face)
        {
            ISurface surf = face.IGetSurface();

            if (surf.IsCylinder())
            {
                double[] uvBounds = face.GetUVBounds() as double[];

                double[] evalData = surf.Evaluate((uvBounds[1] - uvBounds[0]) / 2, (uvBounds[3] - uvBounds[2]) / 2, 1, 1) as double[];

                double[] pt = new double[] { evalData[0], evalData[1], evalData[2] };

                int sense = face.FaceInSurfaceSense() ? 1 : -1;

                double[] norm = new double[] { evalData[evalData.Length - 3] * sense, evalData[evalData.Length - 2] * sense, evalData[evalData.Length - 1] * sense };

                double[] cylParams = surf.CylinderParams as double[];

                double[] orig = new double[] { cylParams[0], cylParams[1], cylParams[2] };

                double[] dir = new double[] { pt[0] - orig[0], pt[1] - orig[1], pt[2] - orig[2] };

                IMathUtility mathUtils = swApp.IGetMathUtility();

                IMathVector dirVec = mathUtils.CreateVector(dir) as IMathVector;
                IMathVector normVec = mathUtils.CreateVector(norm) as IMathVector;

                return GetAngle(dirVec, normVec) < Math.PI / 2;
            }
            else
            {
                throw new NotSupportedException("Only cylindrical face is supported");
            }
        }

        private double GetAngle(IMathVector vec1, IMathVector vec2)
        {
            return Math.Acos(vec1.Dot(vec2) / (vec1.GetLength() * vec2.GetLength()));
        }

        public SldWorks swApp;
    }
}