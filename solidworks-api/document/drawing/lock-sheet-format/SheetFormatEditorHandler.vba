Dim WithEvents swApp As SldWorks.SldWorks

Private Sub Class_Initialize()
    Set swApp = Application.SldWorks
End Sub

Private Function swApp_CommandOpenPreNotify(ByVal Command As Long, ByVal UserCommand As Long) As Long
    
    Const swCommands_Edit_Template As Long = 1501
    
    If Command = swCommands_Edit_Template Then
        Dim cancel As Boolean
        cancel = True
        
        If LOCK_WITH_PASSWORD Then
            
            Dim pwd As String
            PasswordBox.Message = "Sheet format editing is locked. Please enter password to unlock"
            PasswordBox.ShowDialog
            pwd = PasswordBox.Password
            
            If pwd <> "" Then
                If pwd = Password Then
                    cancel = False
                Else
                    swApp.SendMsgToUser2 "Password is incorrect", swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
                End If
            End If
        Else
            swApp.SendMsgToUser2 "Sheet format editing is locked", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
        End If
        
        swApp_CommandOpenPreNotify = IIf(cancel, 1, 0)
    End If
    
End Function