Dim WithEvents swApp As SldWorks.SldWorks

Private Sub UserForm_Initialize()
    Set swApp = Application.SldWorks
End Sub

Private Function swApp_CommandOpenPreNotify(ByVal Command As Long, ByVal UserCommand As Long) As Long
    lstLog.AddItem "Command: " & Command & "; User Command:" & UserCommand
End Function