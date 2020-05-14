    '...
    AddHandler doc.Access3rdPartyData, AddressOf OnAccess3rdPartyData
    '...
Private Sub OnAccess3rdPartyData(ByVal docHandler As DocumentHandler, ByVal state As Access3rdPartyDataState_e)
    Const STREAM_NAME = "CodeStackStream"

    Select Case state
        Case Access3rdPartyDataState_e.StreamWrite

            Using streamHandler = docHandler.Model.Access3rdPartyStream(STREAM_NAME, True)

                Using str = streamHandler.Stream
                    Dim xmlSer = New XmlSerializer(GetType(String()))
                    xmlSer.Serialize(str, New String() {"A", "B"})
                End Using
            End Using
    End Select
End Sub