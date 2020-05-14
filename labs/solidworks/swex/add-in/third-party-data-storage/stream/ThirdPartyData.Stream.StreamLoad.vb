Private Const STREAM_NAME As String = "CodeStackStream"

Public Class StreamData
    Public Property Prp1 As String
    Public Property Prp2 As Double
End Class

Private m_StreamData As StreamData

Private Sub LoadFromStream(ByVal model As IModelDoc2)
    Using streamHandler = model.Access3rdPartyStream(STREAM_NAME, False)

        If streamHandler.Stream IsNot Nothing Then

            Using str = streamHandler.Stream
                Dim xmlSer = New XmlSerializer(GetType(StreamData))
                m_StreamData = TryCast(xmlSer.Deserialize(str), StreamData)
            End Using
        End If
    End Using
End Sub