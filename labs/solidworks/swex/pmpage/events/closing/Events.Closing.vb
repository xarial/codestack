Public Class DataModel
    Public Property Text As String
End Class

Private m_Data As DataModel
Private m_Page As PropertyManagerPageEx(Of MyPMPageHandler, DataModel)

Public Overrides Function OnConnect() As Boolean
    m_Data = New DataModel
    m_Page = New PropertyManagerPageEx(Of MyPMPageHandler, DataModel)(App)

    AddHandler m_Page.Handler.Closing, AddressOf OnClosing
    Return True
End Function

Private Sub OnClosing(ByVal reason As swPropertyManagerPageCloseReasons_e, ByVal arg As ClosingArg)
    If reason = swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay Then

        If String.IsNullOrEmpty(m_Data.Text) Then
            arg.Cancel = True
            arg.ErrorTitle = "Insert Note Error"
            arg.ErrorMessage = "Please specify the note text"
        End If
    End If
End Sub