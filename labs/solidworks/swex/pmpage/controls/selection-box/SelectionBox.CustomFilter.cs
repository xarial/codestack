public class SelectionBoxCustomSelectionFilterDataModel
{
    public class DataGroup
    {
        [SelectionBox(typeof(PlanarFaceFilter), swSelectType_e.swSelFACES)] //setting the standard filter to faces and custom filter to only filter planar faces
        public IFace2 PlanarFace { get; set; }
    }

    public class PlanarFaceFilter : SelectionCustomFilter<IFace2>
    {
        protected override bool Filter(IPropertyManagerPageControlEx selBox, IFace2 selection, swSelectType_e selType, ref string itemText)
        {
            itemText = "Planar Face";
            return selection.IGetSurface().IsPlane(); //validating the selection and only allowing planar face
        }
    }
}