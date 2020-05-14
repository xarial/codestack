Public Event ModelClosed(path As String)
Public Event ModelClosing(path As String)
                         
Dim WithEvents swApp As SldWorks.SldWorks

Dim WithEvents swPart As SldWorks.PartDoc
Dim WithEvents swAssy As SldWorks.AssemblyDoc
Dim WithEvents swDraw As SldWorks.DrawingDoc

Dim IsClosing As Boolean
Dim ModelPath As String

Sub WatchModelClose(model As SldWorks.ModelDoc2)
    
    IsClosing = False
        
    If Not model Is Nothing Then
        Select Case model.GetType()
            Case swDocumentTypes_e.swDocPART
                Set swPart = model
            Case swDocumentTypes_e.swDocASSEMBLY
                Set swAssy = model
            Case swDocumentTypes_e.swDocDRAWING
                Set swDraw = model
        End Select
    Else
        Err.Raise vbError, "", "Model is null"
    End If
    
    ModelPath = model.GetPathName()
    
End Sub

Private Sub Class_Initialize()
    Set swApp = Application.SldWorks
End Sub

Private Function swApp_OnIdleNotify() As Long
    
    If IsClosing Then
        IsClosing = False
        RaiseEvent ModelClosed(ModelPath)
    End If
    
    swApp_OnIdleNotify = 0
    
End Function

Private Function swPart_DestroyNotify2(ByVal DestroyType As Long) As Long
    swPart_DestroyNotify2 = HandleDestroyEvent(DestroyType)
End Function

Private Function swAssy_DestroyNotify2(ByVal DestroyType As Long) As Long
    swAssy_DestroyNotify2 = HandleDestroyEvent(DestroyType)
End Function

Private Function swDraw_DestroyNotify2(ByVal DestroyType As Long) As Long
    swDraw_DestroyNotify2 = HandleDestroyEvent(DestroyType)
End Function

Function HandleDestroyEvent(ByVal DestroyType As Long) As Long
    
    If DestroyType = swDestroyNotifyType_e.swDestroyNotifyDestroy Then
        
        RaiseEvent ModelClosing(ModelPath)
        IsClosing = True
                
    End If
    
    HandleDestroyEvent = 0
    
End Function