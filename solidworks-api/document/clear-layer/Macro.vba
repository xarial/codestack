Const LAYER_NAME As String = "MY LAYER"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swLayerMgr As SldWorks.LayerMgr
    
    Set swLayerMgr = swModel.GetLayerManager
    
    Dim swLayer As SldWorks.layer
    Set swLayer = swLayerMgr.GetLayer(LAYER_NAME)
    
    Dim swLayerItems() As Object
    
    AddItems swLayer, swLayerItemsOption_Annotations, swLayerItems
    AddItems swLayer, swLayerItemsOption_SketchBlockInstance, swLayerItems
    AddItems swLayer, swLayerItemsOption_SketchHatch, swLayerItems
    AddItems swLayer, swLayerItemsOption_SketchPoint, swLayerItems
    AddItems swLayer, swLayerItemsOption_SketchSegments, swLayerItems
    
    If swModel.Extension.MultiSelect(swLayerItems, False, Nothing) = UBound(swLayerItems) + 1 Then
        If False = swModel.Extension.DeleteSelection2(swDeleteSelectionOptions_e.swDelete_Absorbed) Then
            Err.Raise vbError, "", "Failed to delete entities"
        End If
    Else
        Err.Raise vbError, "", "Failed to select items on layer"
    End If
    
End Sub

Sub AddItems(layer As SldWorks.layer, itemsType As swLayerItemsOption_e, ByRef layerItems() As Object)
    
    Dim vItems As Variant
    vItems = layer.GetItems(itemsType)
    
    If Not IsEmpty(vItems) Then

        If (Not layerItems) = -1 Then
            ReDim layerItems(UBound(vItems))
        Else
            ReDim Preserve layerItems(UBound(layerItems) + UBound(vItems) + 1)
        End If
        
        Dim i As Integer
        
        For i = 0 To UBound(vItems)
            Set layerItems(UBound(layerItems) - i) = vItems(i)
        Next
        
    End If
    
End Sub