Public Class ButtonDataModel

    Public ReadOnly Property Button As Action
        Get
            Return AddressOf OnButtonClick
        End Get
    End Property

    Private Sub OnButtonClick()
        'TODO: Handle Button click
    End Sub

End Class
