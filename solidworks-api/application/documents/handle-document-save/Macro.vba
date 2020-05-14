Dim swFileSaveHandler As SaveEventsHandler

Sub main()
    
    Set swFileSaveHandler = New SaveEventsHandler
    
    While True
        DoEvents
    Wend
    
End Sub

Sub OnSaveDocument(Optional dummy As Variant = Empty)
    'TODO: place the code here to run whn document is saved
    MsgBox "Saved"
End Sub