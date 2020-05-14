Public Class SelectionBoxCustomSelectionFilterDataModel

    Public Class DataGroup
        <SelectionBox(GetType(PlanarFaceFilter), swSelectType_e.swSelFACES)>
        Public Property PlanarFace As IFace2
    End Class

    Public Class PlanarFaceFilter
        Inherits SelectionCustomFilter(Of IFace2)

        Protected Overrides Function Filter(ByVal selBox As IPropertyManagerPageControlEx, ByVal selection As IFace2, ByVal selType As swSelectType_e, ByRef itemText As String) As Boolean
            itemText = "Planar Face"
            Return selection.IGetSurface().IsPlane()
        End Function
    End Class

End Class