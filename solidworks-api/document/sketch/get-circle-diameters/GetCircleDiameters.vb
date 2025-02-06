Function GetCircleDiameters(model As IModelDoc2, sketchName As String) As Double()

    Dim diams As New List(Of Double)

    Dim sketchFeat As IFeature

    If TypeOf model Is IPartDoc Then
        sketchFeat = TryCast(model, IPartDoc).FeatureByName(sketchName)
    ElseIf TypeOf model Is IAssemblyDoc Then
        sketchFeat = TryCast(model, IAssemblyDoc).FeatureByName(sketchName)
    Else
        Throw New NotSupportedException()
    End If

    Dim sketch As ISketch = TryCast(sketchFeat.GetSpecificFeature2, ISketch)

    Dim skSegs As Object()

    skSegs = TryCast(sketch.GetSketchSegments(), Object())

    If skSegs IsNot Nothing Then

        For Each skSeg As SketchSegment In skSegs
            If skSeg.GetType() = swSketchSegments_e.swSketchARC Then
                Dim skArc As ISketchArc
                skArc = skSeg
                diams.Add(skArc.GetRadius() * 2)
            End If
        Next

    End If

    Return diams.ToArray()

End Function