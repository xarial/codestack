Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Windows.Forms
Imports System.Text
Imports System

Partial Class CodeStack

    Const ARG_NAME As String = "__SwMacroArgs__"

    Public Sub Main()
        SetArgument("Argument from VB.NET macro")
        Dim err As Integer
        If Not swApp.RunMacro2("D:\Macros\GetArgumentMacro.swp", "Macro1", "main", CInt(swRunMacroOption_e.swRunMacroUnloadAfterRun), err) Then
            swApp.SendMsgToUser(String.Format("Failed to run macro. Error code: {0}", err))
        End If
    End Sub

    Private Shared Sub SetArgument(ByVal arg As String)
        Using stream As MemoryStream = New MemoryStream(Encoding.UTF8.GetBytes(arg))
            Clipboard.SetData(ARG_NAME, stream)
        End Using
    End Sub

    Public swApp As SldWorks

End Class
