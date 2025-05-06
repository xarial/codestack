Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim vAppearances As Variant
    
    vAppearances = swModel.Extension.GetRenderMaterials2(swDisplayStateOpts_e.swAllDisplayState, Empty)
    
    If Not IsEmpty(vAppearances) Then
    
        Dim i As Integer
        
        For i = 0 To UBound(vAppearances)
            Dim swAppearance As SldWorks.RenderMaterial
            Set swAppearance = vAppearances(i)
            
            If HasFace(swAppearance) Then
                Dim matId1(0) As Long
                Dim matId2(0) As Long
                swAppearance.GetMaterialIds matId1(0), matId2(0)
                                
                swModel.Extension.DeleteDisplayStateSpecificRenderMaterial matId1, matId2
            End If
            
        Next
    
    End If
    
End Sub

Function HasFace(appearance As SldWorks.RenderMaterial) As Boolean
    
    Dim vEnts As Variant
    vEnts = appearance.GetEntities
    
    If Not IsEmpty(vEnts) Then
        Dim i As Integer
        For i = 0 To UBound(vEnts)
            If TypeOf vEnts(i) Is SldWorks.Face2 Then
                HasFace = True
                Exit Function
            End If
        Next
    End If
    
    HasFace = False
    
End Function
