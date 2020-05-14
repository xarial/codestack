Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Const HIDE_ALL_SKETCHES As Boolean = False 'True to hide all sketches, False to show all sketches

Const SKETCH_FEAT_TYPE_NAME As String = "ProfileFeature"
Const SKETCH_3D_FEAT_TYPE_NAME As String = "3DProfileFeature"

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then

        Dim swFeat As SldWorks.Feature
    
        Set swFeat = swModel.FirstFeature
        
        While Not swFeat Is Nothing
        
            If swFeat.GetTypeName2 = SKETCH_FEAT_TYPE_NAME Or _
                swFeat.GetTypeName2 = SKETCH_3D_FEAT_TYPE_NAME Then
                
                swFeat.Select2 False, -1
                
                If HIDE_ALL_SKETCHES Then
                    swModel.BlankSketch
                Else
                    swModel.UnblankSketch
                End If
                
            End If
            
            Set swFeat = swFeat.GetNextFeature
            
        Wend
    
    Else
        MsgBox "Please open part or assembly"
    End If
    
End Sub