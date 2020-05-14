Imports CodeStack.SwEx.AddIn
Imports CodeStack.SwEx.AddIn.Attributes
Imports System.Runtime.InteropServices

<ComVisible(True), Guid("A28F5BB7-E468-48B6-9BBD-9E7A31FF8CC8")>
<AutoRegister("OpenGL Box Tetrahedron")>
Public Class AddIn
	Inherits SwAddInEx

	Public Overrides Function OnConnect() As Boolean
		CreateDocumentsHandler(Of OpenGlDocumentHandler)()
		Return True
	End Function
End Class