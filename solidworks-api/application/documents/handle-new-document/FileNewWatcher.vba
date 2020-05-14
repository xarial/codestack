Dim WithEvents swApp As SldWorks.SldWorks

Private Sub Class_Initialize()
    Set swApp = Application.SldWorks
End Sub

Private Function swApp_FileNewNotify2(ByVal NewDoc As Object, ByVal DocType As Long, ByVal TemplateName As String) As Long
    HandlerModule.main NewDoc
End Function