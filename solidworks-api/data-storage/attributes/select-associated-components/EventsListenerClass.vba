Dim swModel As SldWorks.ModelDoc2
Dim WithEvents swAssy As SldWorks.AssemblyDoc
Dim swSelMgr As SldWorks.SelectionMgr

Private Function swAssy_NewSelectionNotify() As Long
    
    Dim swFeat As SldWorks.Feature
    Dim swAtt As SldWorks.Attribute
    Dim swComp As SldWorks.Component2

    Dim i As Integer
    
    i = swSelMgr.GetSelectedObjectCount2(-1)
    
    If i > 0 Then
        
        On Error Resume Next
        
        Set swFeat = swSelMgr.GetSelectedObject6(i, -1)
        
        If Not swFeat Is Nothing Then
        
            If swFeat.GetTypeName2 = "Attribute" Then
            
                Set swAtt = swFeat.GetSpecificFeature2
            
                Set swComp = swAtt.GetComponent()
            
                swComp.Select4 True, Nothing, False
                
            End If
            
        End If
        
    End If
    
    Set swFeat = Nothing
    
End Function

Sub SetAssembly(assy As SldWorks.AssemblyDoc)
        
    Set swAssy = assy
    
    Set swModel = swAssy
        
    Set swSelMgr = swModel.SelectionManager
       
End Sub

