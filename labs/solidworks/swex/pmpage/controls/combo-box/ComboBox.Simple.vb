Imports CodeStack.SwEx.Common.Attributes
Imports CodeStack.SwEx.My.Resources
Imports CodeStack.SwEx.PMPage.Attributes
Imports SolidWorks.Interop.swconst

Public Class ComboBoxDataModel

    Public Enum Options_e
        Option1
        Option2
        Option3
    End Enum

    <ComboBoxOptions(swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_Sorted)>
    Public Property Options As Options_e

End Class
