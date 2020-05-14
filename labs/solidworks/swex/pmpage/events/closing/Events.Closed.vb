Public Class DataModel
    Public Property Text As String
End Class

Private m_Data As DataModel
Private m_Page As PropertyManagerPageEx(Of MyPMPageHandler, DataModel)

Public Overrides Function OnConnect() As Boolean
    m_Data = New DataModel
    m_Page = New PropertyManagerPageEx(Of MyPMPageHandler, DataModel)(App)

    AddHandler m_Page.Handler.Closed, AddressOf OnClosed
    Return True
End Function

Private Sub OnClosed(ByVal reason As swPropertyManagerPageCloseReasons_e)
    If reason = swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay Then
        'TODO: do work
    Else
        'TODO: release resources
    End If
End Sub