Public Enum swPlanes_e
    Front = 1
    Top = 2
    Right = 3
End Enum

Dim swApp As SldWorks.SldWorks
    
Sub main()
               
    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        If swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY Or _
            swModel.GetType() = swDocumentTypes_e.swDocPART Then
            
            SelectPlane swModel, swPlanes_e.Right
            
        Else
            MsgBox "Only assemblies and parts are supported"
        End If
    Else
        MsgBox "Please open part or assembly"
    End If
    
End Sub

Sub SelectPlane(model As SldWorks.ModelDoc2, planeType As swPlanes_e)

    Dim planeIndex As Integer
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = model.FirstFeature

    Do While Not swFeat Is Nothing

        If swFeat.GetTypeName = "RefPlane" Then

            planeIndex = planeIndex + 1
            
            If CInt(planeType) = planeIndex Then

                swFeat.Select2 False, 0

                Exit Sub

            End If

        End If
    
        Set swFeat = swFeat.GetNextFeature

    Loop
    
End Sub