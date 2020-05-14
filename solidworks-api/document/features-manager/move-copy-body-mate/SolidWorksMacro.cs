using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace MoveBodyMate
{
    public partial class SolidWorksMacro
    {
        public void Main()
        {
            try
            {
                var model = swApp.IActiveDoc2;
                var body = model.ISelectionManager.GetSelectedObject6(1, -1) as IBody2;

                if (body == null)
                {
                    throw new NullReferenceException("Select body to move");
                }

                var plane = FindFrontPlane(model);
                var face = FindLargestPlanarFace(body);

                AddCoincidentMate(model, body, plane as IEntity, face as IEntity);
            }
            catch(Exception ex)
            {
                swApp.SendMsgToUser2(ex.Message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
            }
        }

        private static void AddCoincidentMate(ModelDoc2 model, IBody2 body, IEntity mateEnt, IEntity mateEntOnBody)
        {
            var selData = model.ISelectionManager.CreateSelectData();
            selData.Mark = 1;

            if (!body.Select2(false, selData))
            {
                throw new InvalidOperationException("Failed to select body");
            }

            var moveCopyBodyFeat = model.FeatureManager.InsertMoveCopyBody2(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, false, 1);

            var featData = moveCopyBodyFeat.GetDefinition() as IMoveCopyBodyFeatureData;

            if (featData.AccessSelections(model, null))
            {
                int err;

                var mate = featData.AddMate(new DispatchWrapper[]
                {
                    new DispatchWrapper(mateEnt),
                    new DispatchWrapper(mateEntOnBody)
                }, (int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignCLOSEST, 0, 0, out err);

                if (mate == null)
                {
                    throw new NullReferenceException(string.Format("Failed to add mate: {0}", (swAddMateError_e)err));
                }

                if (!moveCopyBodyFeat.ModifyDefinition(featData, model, null))
                {
                    throw new InvalidOperationException("Failed to apply feature changes");
                }
            }
            else
            {
                throw new InvalidOperationException("Failed to access the definition");
            }
        }

        private IFeature FindFrontPlane(IModelDoc2 model)
        {
            var feat = model.IFirstFeature();

            while (feat != null && feat.GetTypeName2() != "RefPlane")
            {
                feat = feat.IGetNextFeature();
            }

            if (feat == null)
            {
                throw new NullReferenceException("Failed to find the front plane");
            }

            return feat;
        }

        private IFace2 FindLargestPlanarFace(IBody2 body)
        {
            var faces = body.GetFaces() as object[];

            if (faces == null)
            {
                throw new NullReferenceException("Body doesn't contain faces");
            }

            var face = faces.Cast<IFace2>()
                .Where(f => f.IGetSurface().IsPlane())
                .OrderBy(f => f.GetArea()).LastOrDefault();

            if (face == null)
            {
                throw new NullReferenceException("There are no planar faces in this body");
            }

            return face;
        }

        public SldWorks swApp;
    }
}

