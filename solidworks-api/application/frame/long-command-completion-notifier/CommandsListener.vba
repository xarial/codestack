Dim WithEvents swApp As SldWorks.SldWorks

Dim IsCommandStarted As Boolean
Dim StartCommand As Long
Dim StartCommandTimeStamp As Date

Public MinimumDelay As Double

Private Sub Class_Initialize()
    Set swApp = Application.SldWorks
End Sub

Private Function swApp_CommandOpenPreNotify(ByVal Command As Long, ByVal UserCommand As Long) As Long
    IsCommandStarted = True
    StartCommand = Command
    StartCommandTimeStamp = Now
    swApp_CommandOpenPreNotify = 0
End Function

Private Function swApp_CommandCloseNotify(ByVal Command As Long, ByVal reason As Long) As Long
    
    If IsCommandStarted And Command = StartCommand Then
    
        IsCommandStarted = False
    
        Dim diff As Integer
        diff = CInt(DateDiff("s", StartCommandTimeStamp, Now))
        
        Debug.Print diff
        
        If diff >= MinimumDelay Then
            Beep
        End If
        
    End If
    
    swApp_CommandCloseNotify = 0
    
End Function