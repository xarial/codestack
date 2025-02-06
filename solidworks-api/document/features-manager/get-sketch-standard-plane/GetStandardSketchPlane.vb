Function GetStandardSketchPlane(app As ISldWorks, model As IModelDoc2, sketchName As String) As String

    Dim sketchFeat As IFeature

    If TypeOf model Is IPartDoc Then
        sketchFeat = TryCast(model, IPartDoc).FeatureByName(sketchName)
    ElseIf TypeOf model Is IAssemblyDoc Then
        sketchFeat = TryCast(model, IAssemblyDoc).FeatureByName(sketchName)
    Else
        Throw New NotSupportedException()
    End If

    If sketchFeat IsNot Nothing Then

        Dim refType As Integer
        Dim refEnt As Object

        Dim sketch As ISketch = TryCast(sketchFeat.GetSpecificFeature2, ISketch)

        refEnt = sketch.GetReferenceEntity(refType)

        If refType = swSelectType_e.swSelDATUMPLANES Then
            Dim feat As IFeature

            feat = model.IFirstFeature

            Dim standardPlaneIndex As Integer = 0

            Do While feat IsNot Nothing And feat.GetTypeName2() <> "OriginProfileFeature"
                If feat.GetTypeName2() = "RefPlane" Then
                    standardPlaneIndex = standardPlaneIndex + 1

                    If app.IsSame(feat.GetSpecificFeature2(), refEnt) = swObjectEquality.swObjectSame Then
                        Select Case standardPlaneIndex
                            Case 1
                                Return "Front"
                            Case 2
                                Return "Top"
                            Case 3
                                Return "Right"
                            Case Else
                                Throw New NotSupportedException("Standard plane is not supported")
                        End Select
                    End If
                End If
                feat = feat.IGetNextFeature
            Loop

            Throw New NotSupportedException("Reference entity is not a standard plane")
        Else
            Throw New NotSupportedException("Reference entity is not a plane")
        End If

    Else
        Throw New NullReferenceException("Failed to find the sketch by name")
    End If

End Function