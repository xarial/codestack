Public Class SelectionBoxListDataModel

    <SelectionBox(swSelectType_e.swSelSOLIDBODIES)>
    Public Property Bodies As List(Of IBody2) = New List(Of IBody2)()

    <SelectionBox(swSelectType_e.swSelEDGES, swSelectType_e.swSelNOTES, swSelectType_e.swSelCOORDSYS)>
    Public Property Dispatches As List(Of Object) = New List(Of Object)()

End Class