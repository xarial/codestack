Const TAG As String = "_CodeStackNote_"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        If Not TagSelectedNote(swModel, TAG) Then
            MsgBox "Failed to add tag to the note"
        End If
        
    Else
        MsgBox "Please open the model"
    End If
    
End Sub

Function TagSelectedNote(model As SldWorks.ModelDoc2, TAG As String) As Boolean
    
    On Error Resume Next
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager
            
    Dim swNote As SldWorks.Note
    
    Set swNote = swSelMgr.GetSelectedObject6(1, -1)
    
    If Not swNote Is Nothing Then
        swNote.TagName = TAG
        TagSelectedNote = True
        Exit Function
    Else
        MsgBox "Please select note to add tag to"
    End If
    
    TagSelectedNote = False
    
End Function