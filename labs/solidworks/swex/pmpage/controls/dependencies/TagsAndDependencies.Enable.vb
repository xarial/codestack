Public Class DataModelEnable

    <ControlTag(NameOf(Enable))>
    Public Property Enable As Boolean

    <SelectionBox(swSelectType_e.swSelFACES)>
    <DependentOn(GetType(EnableDepHandler), NameOf(Enable))>
    Public Property Selection As IEntity

End Class

Public Class EnableDepHandler
    Inherits DependencyHandler

    Protected Overrides Sub UpdateControlState(ByVal control As IPropertyManagerPageControlEx, ByVal parents As IPropertyManagerPageControlEx())
        control.Enabled = CBool(parents.First().GetValue())
    End Sub
End Class