Public Overrides Function OnConnect() As Boolean
    AddContextMenu(Of CommandsD_e)(AddressOf OnCommandsDContextMenuClick)
    AddContextMenu(Of CommandsE_e)(AddressOf OnCommandsEContextMenuClick, swSelectType_e.swSelFACES)
    Return True
End Function

Private Sub OnCommandsDContextMenuClick(ByVal cmd As CommandsD_e)
    'TODO: handle the context menu click
End Sub

Private Sub OnCommandsEContextMenuClick(ByVal cmd As CommandsE_e)
    'TODO: handle the context menu click
End Sub