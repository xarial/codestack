Const MIN_DELAY As Integer = 5

Dim swCmdsListener As CommandsListener

Sub main()

    Set swCmdsListener = New CommandsListener
    swCmdsListener.MinimumDelay = MIN_DELAY
    
End Sub