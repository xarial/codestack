Dim WithEvents swApp As SldWorks.SldWorks

Dim swCutListFeat As SldWorks.Feature
Dim swModel As SldWorks.ModelDoc2
Dim LinkedProperties As Variant

Private Sub Class_Initialize()
    LinkedProperties = Array("DESCRIPTION", "LENGTH", "QUANTITY")
End Sub

Sub Init(app As SldWorks.SldWorks, model As SldWorks.ModelDoc2, cutListFeat As SldWorks.Feature)
    
    Set swApp = app
    
    Set swModel = model
    Set swCutListFeat = cutListFeat
    
End Sub

Private Function swApp_OnIdleNotify() As Long
    CopyProperties
    Set swApp = Nothing 'unsubscribe from the event
End Function

Sub CopyProperties()
    
    Dim i As Integer
    
    Dim swSrcPrpMgr As SldWorks.CustomPropertyManager
    Set swSrcPrpMgr = swCutListFeat.CustomPropertyManager
    
    Dim swDestPrpMgr As SldWorks.CustomPropertyManager
    Set swDestPrpMgr = swModel.Extension.CustomPropertyManager("")
    
    For i = 0 To UBound(LinkedProperties)
    
        Dim prpName As String
        prpName = CStr(LinkedProperties(i))
        
        Dim prpVal As String

        swSrcPrpMgr.Get2 prpName, "", prpVal
        
        swDestPrpMgr.Add2 prpName, swCustomInfoType_e.swCustomInfoText, prpVal
        swDestPrpMgr.Set prpName, prpVal
        
    Next
    
End Sub