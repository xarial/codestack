Public Class TaskPaneControl
    Inherits UserControl
End Class
Public Enum TaskPaneCommands_e
    Command1
End Enum
    '...
    Dim ctrl As TaskPaneControl = Nothing
    Dim taskPaneView = CreateTaskPane(Of TaskPaneControl, TaskPaneCommands_e)(AddressOf OnTaskPaneCommandClick, ctrl)
    '...
Private Sub OnTaskPaneCommandClick(ByVal cmd As TaskPaneCommands_e)
    Select Case cmd
        Case TaskPaneCommands_e.Command1
    End Select
End Sub