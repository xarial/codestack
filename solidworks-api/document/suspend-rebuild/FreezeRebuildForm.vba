Const CANCEL_REGEN As Long = 1

Dim swApp As SldWorks.SldWorks

Dim WithEvents swPart As SldWorks.PartDoc
Dim WithEvents swAssy As SldWorks.AssemblyDoc
Dim WithEvents swDraw As SldWorks.DrawingDoc

Private Sub btnExit_Click()
    End
End Sub

Private Sub UserForm_Initialize()
    
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        Select Case swModel.GetType()
            Case swDocumentTypes_e.swDocPART
                Set swPart = swModel
            Case swDocumentTypes_e.swDocASSEMBLY
                Set swAssy = swModel
            Case swDocumentTypes_e.swDocDRAWING
                Set swDraw = swModel
        End Select
            
    Else
        MsgBox "Please open the model"
        End
    End If
    
End Sub

Private Function swAssy_RegenNotify() As Long
    swAssy_RegenNotify = CANCEL_REGEN
End Function

Private Function swDraw_RegenNotify() As Long
    swDraw_RegenNotify = CANCEL_REGEN
End Function

Private Function swPart_RegenNotify() As Long
    swPart_RegenNotify = CANCEL_REGEN
End Function