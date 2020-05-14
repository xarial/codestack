Sub main()
    
    Dim cmdsMgr As CommandRunManager
    Set cmdsMgr = New CommandRunManager
    
    If cmdsMgr.RunCommand(swCommands_Extrude) Then
        MsgBox "Command Completed"
    Else
        MsgBox "Command Cancelled"
    End If
    
End Sub
