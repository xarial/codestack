Public Password As String

Public Property Let Message(msg As String)
     lblMessage.Caption = msg
End Property

Public Sub ShowDialog()
    Password = ""
    txtPassword.Text = ""
    Show vbModal
End Sub

Private Sub btnOk_Click()
    
    Password = txtPassword.Text
    Me.Hide
    
End Sub