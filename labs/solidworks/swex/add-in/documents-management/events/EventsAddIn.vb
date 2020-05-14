Imports CodeStack.SwEx.AddIn
Imports CodeStack.SwEx.AddIn.Attributes
Imports CodeStack.SwEx.AddIn.Base
Imports CodeStack.SwEx.AddIn.Core
Imports CodeStack.SwEx.AddIn.Delegates
Imports CodeStack.SwEx.AddIn.Enums
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports System
Imports System.Runtime.InteropServices

Namespace CodeStack.SwEx
    <AutoRegister>
    <ComVisible(True), Guid("E6BE0C5A-8B24-46B4-98F9-BEC4100BC709")>
    Public Class EventsAddIn
        Inherits SwAddInEx

        Private m_DocHandlerGeneric As IDocumentsHandler(Of DocumentHandler)

        Public Overrides Function OnConnect() As Boolean
            m_DocHandlerGeneric = CreateDocumentsHandler()
            AddHandler m_DocHandlerGeneric.HandlerCreated, AddressOf OnHandlerCreated
            Return True
        End Function

        Private Sub OnHandlerCreated(ByVal doc As DocumentHandler)
            AddHandler doc.Initialized, AddressOf OnInitialized
            AddHandler doc.Activated, AddressOf OnActivated
            AddHandler doc.ConfigurationChange, AddressOf OnConfigurationOrSheetChanged
            AddHandler doc.CustomPropertyModify, AddressOf OnCustomPropertyModified
            AddHandler doc.Access3rdPartyData, AddressOf OnAccess3rdPartyData
            AddHandler doc.ItemModify, AddressOf OnItemModified
            AddHandler doc.Save, AddressOf OnSave
            AddHandler doc.Selection, AddressOf OnSelection
            AddHandler doc.Rebuild, AddressOf OnRebuild
            AddHandler doc.DimensionChange, AddressOf OnDimensionChange
            AddHandler doc.Destroyed, AddressOf OnDestroyed
        End Sub

        Private Sub OnDestroyed(ByVal handler As DocumentHandler)
            RemoveHandler handler.Initialized, AddressOf OnInitialized
            RemoveHandler handler.Activated, AddressOf OnActivated
            RemoveHandler handler.ConfigurationChange, AddressOf OnConfigurationOrSheetChanged
            RemoveHandler handler.CustomPropertyModify, AddressOf OnCustomPropertyModified
            RemoveHandler handler.ItemModify, AddressOf OnItemModified
            RemoveHandler handler.Save, AddressOf OnSave
            RemoveHandler handler.Selection, AddressOf OnSelection
            RemoveHandler handler.Rebuild, AddressOf OnRebuild
            RemoveHandler handler.DimensionChange, AddressOf OnDimensionChange
            RemoveHandler handler.Destroyed, AddressOf OnDestroyed
            Logger.Log($"'{handler.Model.GetTitle()}' destroyed")
        End Sub

        Public Overrides Function OnDisconnect() As Boolean
            RemoveHandler m_DocHandlerGeneric.HandlerCreated, AddressOf OnHandlerCreated
            Return True
        End Function

    End Class
End Namespace
