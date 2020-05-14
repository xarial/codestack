using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace CodeStack
{
    public partial class SolidWorksMacro
    {
        public enum FaceShellType_e
        {
            Open = 0,
            Internal = 1,
            External = 2
        }

        public void Main()
        {
            IModelDoc2 model = swApp.IActiveDoc2;

            if (model != null)
            {
                SelectionMgr selMgr = model.ISelectionManager;

                IBody2 body = selMgr.GetSelectedObject6(1, -1) as IBody2;

                if (body != null)
                {
                    if (body.GetType() == (int)swBodyType_e.swSheetBody)
                    {
                        if (IsOpenGeometry(body))
                        {
                            swApp.SendMsgToUser("Selected body is an open geometry");
                        }
                        else
                        {
                            swApp.SendMsgToUser("Selected body is not an open geometry");
                        }
                    }
                    else
                    {
                        swApp.SendMsgToUser("Selected body is not a sheet body");
                    }
                }
                else
                {
                    swApp.SendMsgToUser("Please select sheet body");
                }
            }
            else
            {
                swApp.SendMsgToUser("Please open model");
            }

            return;
        }

        private static bool IsOpenGeometry(IBody2 body)
        {
            object[] faces = body.GetFaces() as object[];

            if (faces != null)
            {
                foreach (IFace2 face in faces)
                {
                    FaceShellType_e shellType = (FaceShellType_e)face.GetShellType();

                    if (shellType == FaceShellType_e.Open)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public SldWorks swApp;
    }
}
