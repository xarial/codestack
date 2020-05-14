Dim WithEvents swPart As SldWorks.PartDoc
Dim WithEvents swAssy As SldWorks.AssemblyDoc
Dim WithEvents swDraw As SldWorks.DrawingDoc

Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr

Dim swSelFilter As Integer

Sub WaitForSelection(model As SldWorks.ModelDoc2, selFilter As Integer)
        
    Set swModel = model
    swSelFilter = selFilter
            
    Set swSelMgr = swModel.SelectionManager
            
    If TypeOf model Is SldWorks.PartDoc Then
        Set swPart = model
    ElseIf TypeOf model Is SldWorks.AssemblyDoc Then
        Set swAssy = model
    ElseIf TypeOf model Is SldWorks.DrawingDoc Then
        Set swDraw = model
    End If
    
End Sub

Private Function swPart_NewSelectionNotify() As Long
    HandleSelection
End Function

Private Function swAssy_NewSelectionNotify() As Long
    HandleSelection
End Function

Private Function swDraw_NewSelectionNotify() As Long
    HandleSelection
End Function

Sub HandleSelection()
    
    Dim selCount As Integer
    selCount = swSelMgr.GetSelectedObjectCount2(-1)
    
    If selCount > 0 Then
        If swSelMgr.GetSelectedObjectType3(selCount, -1) = swSelFilter Then
            Dim swObject As Object
            Set swObject = swSelMgr.GetSelectedObject6(selCount, -1)
            Stop
        End If
    End If
End Sub