<AutoRegister("My C# SOLIDWORKS Add-In", "Sample SOLIDWORKS add-in in VB.NET", True)>
<ComVisible(True), Guid("31E2C0F0-B68D-44C4-AB15-4CC7B56B6C16")>
Public Class SampleAddIn
    Inherits SwAddInEx

    Public Overrides Function OnConnect() As Boolean
        Return True
    End Function

End Class