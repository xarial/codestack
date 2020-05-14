Public Class DataModel
    Public Property Text As String
End Class

Private m_Data As DataModel
Private m_Page As PropertyManagerPageEx(Of MyPMPageHandler, DataModel)

Public Overrides Function OnConnect() As Boolean
    m_Data = New DataModel
    m_Page = New PropertyManagerPageEx(Of MyPMPageHandler, DataModel)(App)

    AddHandler m_Page.Handler.DataChanged, AddressOf OnDataChanged
    Return True
End Function

Private Sub OnDataChanged()
    Dim text = m_Data.Text
    'TODO: handle the data changing, e.g. update preview
End Sub