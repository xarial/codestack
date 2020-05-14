Public Class DataModelIgnore
	Public Property Text As String

	<IgnoreBinding>
	Public Property CalculatedField As Integer 'control will not be generated for this field
End Class