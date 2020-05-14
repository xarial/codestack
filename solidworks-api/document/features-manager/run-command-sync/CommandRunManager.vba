Dim WithEvents swApp As SldWorks.SldWorks

Dim CurrentCommandId As Long
Dim IsCommandCompleted As Boolean
Dim CloseReason As Long

Private Sub Class_Initialize()
    
    Set swApp = Application.SldWorks
    
End Sub

Function RunCommand(cmd As swCommands_e) As Boolean
    
    IsCommandCompleted = False
    CurrentCommandId = cmd
    swApp.RunCommand cmd, ""
    
    While Not IsCommandCompleted
        DoEvents
    Wend
    
    RunCommand = CloseReason = swCommands_e.swCommands_PmOK
    
End Function

Private Function swApp_CommandCloseNotify(ByVal Command As Long, ByVal reason As Long) As Long
    
    If CurrentCommandId <> -1 Then
    
        If Command = CurrentCommandId Then
            CurrentCommandId = -1
            IsCommandCompleted = True
            CloseReason = reason
        End If
    
    End If
    
End Function