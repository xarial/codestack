Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swpublished
Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.InteropServices

<ComVisible(True)>
<Guid("799707B3-1632-469F-B294-EC05A5FBFFC8")>
<DisplayName("Sample Add-In")>
<Description("Sample 'Hello World' SOLIDWORKS add-in")>
Public Class MySampleAddin
    Implements ISwAddin

    Private Const ADDIN_KEY_TEMPLATE As String = "SOFTWARE\SolidWorks\Addins\{{{0}}}"
    Private Const ADDIN_STARTUP_KEY_TEMPLATE As String = "Software\SolidWorks\AddInsStartup\{{{0}}}"
    Private Const ADD_IN_TITLE_REG_KEY_NAME As String = "Title"
    Private Const ADD_IN_DESCRIPTION_REG_KEY_NAME As String = "Description"

#Region "Registration"

    <ComRegisterFunction>
    Public Shared Sub RegisterFunction(ByVal t As Type)
        Try
            Dim addInTitle = ""
            Dim loadAtStartup = True
            Dim addInDesc = ""
            Dim dispNameAtt = t.GetCustomAttributes(False).OfType(Of DisplayNameAttribute)().FirstOrDefault()

            If dispNameAtt IsNot Nothing Then
                addInTitle = dispNameAtt.DisplayName
            Else
                addInTitle = t.ToString()
            End If

            Dim descAtt = t.GetCustomAttributes(False).OfType(Of DescriptionAttribute)().FirstOrDefault()

            If descAtt IsNot Nothing Then
                addInDesc = descAtt.Description
            Else
                addInDesc = t.ToString()
            End If

            Dim addInkey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(String.Format(ADDIN_KEY_TEMPLATE, t.GUID))
            addInkey.SetValue(Nothing, 0)
            addInkey.SetValue(ADD_IN_TITLE_REG_KEY_NAME, addInDesc)
            addInkey.SetValue(ADD_IN_DESCRIPTION_REG_KEY_NAME, addInTitle)
            Dim addInStartupkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(String.Format(ADDIN_STARTUP_KEY_TEMPLATE, t.GUID))
            addInStartupkey.SetValue(Nothing, Convert.ToInt32(loadAtStartup), Microsoft.Win32.RegistryValueKind.DWord)
        Catch ex As Exception
            Console.WriteLine("Error while registering the addin: " & ex.Message)
        End Try

    End Sub

    <ComUnregisterFunction>
    Public Shared Sub UnregisterFunction(ByVal t As Type)
        Try
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKey(String.Format(ADDIN_KEY_TEMPLATE, t.GUID))
            Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(String.Format(ADDIN_STARTUP_KEY_TEMPLATE, t.GUID))
        Catch e As Exception
            Console.WriteLine("Error while unregistering the addin: " & e.Message)
        End Try
    End Sub
#End Region

    Private m_App As ISldWorks

    Public Function ConnectToSW(ByVal ThisSW As Object, ByVal Cookie As Integer) As Boolean Implements ISwAddin.ConnectToSW
        m_App = TryCast(ThisSW, ISldWorks)
        m_App.SendMsgToUser("Hello World!")
        Return True
    End Function

    Public Function DisconnectFromSW() As Boolean Implements ISwAddin.DisconnectFromSW
        Return True
    End Function

End Class
