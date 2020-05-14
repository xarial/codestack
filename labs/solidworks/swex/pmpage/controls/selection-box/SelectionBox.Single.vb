Public Class SelectionBoxDataModel

    <SelectionBox(swSelectType_e.swSelSOLIDBODIES)>
    Public Property Body As IBody2

    <SelectionBox(swSelectType_e.swSelEDGES, swSelectType_e.swSelNOTES, swSelectType_e.swSelCOORDSYS)>
    Public Property Dispatch As Object

End Class