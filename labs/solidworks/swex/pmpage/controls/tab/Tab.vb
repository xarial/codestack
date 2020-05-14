Imports CodeStack.SwEx.Common.Attributes
Imports CodeStack.SwEx.My.Resources
Imports CodeStack.SwEx.PMPage.Attributes

Public Class TabDataModel

	<Tab>
	<Icon(GetType(Resources), NameOf(Resources.OffsetImage))>
	Public Class TabControl1
		Public Property Field1 As String
	End Class

	Public Property Tab1 As TabControl1

End Class
