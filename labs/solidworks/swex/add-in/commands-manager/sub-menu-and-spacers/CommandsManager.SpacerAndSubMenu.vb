
<Title("AddInEx Commands")>
Public Enum Commands_e

    Command1

    <CommandSpacer>
    Command2

End Enum

<Title("Sub Menu Commands")>
<CommandGroupInfo(GetType(Commands_e))>
Public Enum SubCommands_e
    SubCommand1
    SubCommand2
End Enum

Public Overrides Function OnConnect() As Boolean
    AddCommandGroup(Of Commands_e)(AddressOf OnButtonClick)
    AddCommandGroup(Of SubCommands_e)(AddressOf OnButtonClick)
    Return True
End Function

Private Sub OnButtonClick(ByVal cmd As Commands_e)
End Sub

Private Sub OnButtonClick(ByVal cmd As SubCommands_e)
End Sub