Public Class MyDocHandler
    Implements IDocumentHandler

    Private m_Model As IModelDoc2

    Public Sub Init(ByVal app As ISldWorks, ByVal model As IModelDoc2) Implements IDocumentHandler.Init
        If TypeOf model Is PartDoc Then
            m_Model = model
            AddHandler(TryCast(m_Model, PartDoc)).AddItemNotify, AddressOf OnAddItemNotify
        End If
    End Sub

    Private Function OnAddItemNotify(ByVal EntityType As Integer, ByVal itemName As String) As Integer
        Return 0
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If TypeOf m_Model Is PartDoc Then
            RemoveHandler(TryCast(m_Model, PartDoc)).AddItemNotify, AddressOf OnAddItemNotify
        End If
    End Sub

End Class
