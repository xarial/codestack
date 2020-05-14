Private Const STORAGE_NAME As String = "CodeStackStorage"
Private Const STREAM1_NAME As String = "CodeStackStream1"
Private Const STREAM2_NAME As String = "CodeStackStream2"
Private Const SUB_STORAGE_NAME As String = "CodeStackSubStorage"

Public Class StorageStreamData
    Public Property Prp3 As Integer
    Public Property Prp4 As Boolean
End Class

Private m_StorageData As StorageStreamData
Private Sub LoadFromStorageStore(ByVal model As IModelDoc2)
    Using storageHandler = model.Access3rdPartyStorageStore(STORAGE_NAME, False)

        If storageHandler.Storage IsNot Nothing Then

            Using str = storageHandler.Storage.TryOpenStream(STREAM1_NAME, False)

                If str IsNot Nothing Then
                    Dim xmlSer = New XmlSerializer(GetType(StorageStreamData))
                    m_StorageData = TryCast(xmlSer.Deserialize(str), StorageStreamData)
                End If
            End Using

            Using subStorage = storageHandler.Storage.TryOpenStorage(SUB_STORAGE_NAME, False)

                If subStorage IsNot Nothing Then

                    Using str = subStorage.TryOpenStream(STREAM2_NAME, False)

                        If str IsNot Nothing Then
                            Dim buffer = New Byte(str.Length - 1) {}
                            str.Read(buffer, 0, buffer.Length)
                            Dim dateStr = Encoding.UTF8.GetString(buffer)
                            Dim [date] = DateTime.Parse(dateStr)
                        End If
                    End Using
                End If
            End Using
        End If
    End Using
End Sub