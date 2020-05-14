Imports CodeStack.SwEx.AddIn
Imports CodeStack.SwEx.AddIn.Attributes
Imports CodeStack.SwEx.AddIn.Base
Imports CodeStack.SwEx.AddIn.Core
Imports CodeStack.SwEx.AddIn.Enums
Imports SolidWorks.Interop.sldworks
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Xml.Serialization

Namespace CodeStack.SwEx
    <AutoRegister>
    <ComVisible(True), Guid("0421C699-E1B1-4D64-A086-D686E88EC311")>
    Public Class ThirdPartyDataAddIn
        Inherits SwAddInEx

        Private m_StorageDocHandler As IDocumentsHandler(Of DocumentHandler)

        Public Overrides Function OnConnect() As Boolean
            m_StorageDocHandler = CreateDocumentsHandler()
            AddHandler m_StorageDocHandler.HandlerCreated, AddressOf OnStorageHandlerCreated
            Return True
        End Function

        Private Sub OnStorageHandlerCreated(ByVal doc As DocumentHandler)
            AddHandler doc.Access3rdPartyData, AddressOf OnAccess3rdPartyStorageStore
        End Sub

        Private Sub OnAccess3rdPartyStorageStore(ByVal docHandler As DocumentHandler, ByVal state As Access3rdPartyDataState_e)
            Select Case state
                Case Access3rdPartyDataState_e.StorageRead
                    LoadFromStorageStore(docHandler.Model)
                Case Access3rdPartyDataState_e.StorageWrite
                    SaveToStorageStore(docHandler.Model)
            End Select
        End Sub
    End Class
End Namespace
