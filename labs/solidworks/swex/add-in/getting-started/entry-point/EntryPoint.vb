Imports CodeStack.SwEx.AddIn
Imports CodeStack.SwEx.AddIn.Attributes
Imports System.Runtime.InteropServices

<AutoRegister("MyAddIn title", "MyAddIn description", True)>
<ComVisible(True), Guid("025F9A68-F2FE-46CF-8BA2-8E19FBCDE9A0")>
Public Class MyAddIn
    Inherits SwAddInEx

    Public Overrides Function OnConnect() As Boolean
        'Initialize the add-in, create menu, load data etc.
        Return True
    End Function

    Public Overrides Function OnDisconnect() As Boolean
        'Dispose the add-in's resources
        Return True
    End Function

End Class
