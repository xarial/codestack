Public Enum Groups_e
    GroupA
    GroupB
    GroupC
End Enum

Public Enum GroupA_e
    GroupA_OptionA
    GroupA_OptionB
    GroupA_OptionC
End Enum

Public Enum GroupB_e
    GroupB_OptionA
    GroupB_OptionB
End Enum

Public Enum GroupC_e
    GroupC_OptionA
    GroupC_OptionB
    GroupC_OptionC
    GroupC_OptionD
End Enum

Public Enum Tags_e
    Groups
End Enum

Public Class DataModelCascading

    <ControlTag(Tags_e.Groups)>
    Public Property Groups As Groups_e

    <DependentOn(GetType(GroupOptionsVisibilityDepHandler), Tags_e.Groups)>
    <ControlTag(Groups_e.GroupA)>
    <OptionBox>
    Public Property GroupA As GroupA_e

    <DependentOn(GetType(GroupOptionsVisibilityDepHandler), Tags_e.Groups)>
    <ControlTag(Groups_e.GroupB)>
    <OptionBox>
    Public Property GroupB As GroupB_e

    <DependentOn(GetType(GroupOptionsVisibilityDepHandler), Tags_e.Groups)>
    <ControlTag(Groups_e.GroupC)>
    <OptionBox>
    Public Property GroupC As GroupC_e

End Class

Public Class GroupOptionsVisibilityDepHandler
    Inherits DependencyHandler

    Protected Overrides Sub UpdateControlState(ByVal control As IPropertyManagerPageControlEx, ByVal parents As IPropertyManagerPageControlEx())
        Dim curGrp = CType(parents.First().GetValue(), Groups_e)
        control.Visible = CType(control.Tag, Groups_e) = curGrp
    End Sub

End Class