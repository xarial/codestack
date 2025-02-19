Const LAYER_NAME As String = "Hatch"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = swModel.SelectionManager
    
    Dim swView As SldWorks.view
    
    Set swView = swSelMgr.GetSelectedObjectsDrawingView2(1, -1)
    
    If Not swView Is Nothing Then
    
        Dim vFaceHatches As Variant
        vFaceHatches = swView.GetFaceHatches
        
        If Not IsEmpty(vFaceHatches) Then
            Dim i As Integer
            
            For i = 0 To UBound(vFaceHatches)
                Dim swFaceHatch As SldWorks.FaceHatch
                Set swFaceHatch = vFaceHatches(i)
                
                swFaceHatch.Layer = LAYER_NAME
            Next
            
        End If
        
        RefreshView swModel, swView
        
    Else
        Err.Raise vbError, "", "Select drawing view"
    End If
    
End Sub

Sub RefreshView(draw As SldWorks.DrawingDoc, swView As SldWorks.view)
    
    If SelectDrawingView(draw, swView) Then
        
        draw.SuppressView
        
        If SelectDrawingView(draw, swView) Then
            draw.UnsuppressView
        End If
        
    End If

End Sub

Function SelectDrawingView(draw As SldWorks.ModelDoc2, view As SldWorks.view) As Boolean
    SelectDrawingView = False <> draw.Extension.SelectByID2(view.Name, "DRAWINGVIEW", 0, 0, 0, False, -1, Nothing, swSelectOption_e.swSelectOptionDefault)
End Function