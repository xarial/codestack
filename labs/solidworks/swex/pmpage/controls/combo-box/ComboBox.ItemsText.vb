Imports CodeStack.SwEx.Common.Attributes
Imports CodeStack.SwEx.My.Resources
Imports CodeStack.SwEx.PMPage.Attributes
Imports SolidWorks.Interop.swconst

Public Class ComboBoxDataModel

    Public Enum OptionsCustomized_e
        <Title("First Option")>
        Option1
        <Title(GetType(Resources), NameOf(Resources.Option2Title))>
        Option2
    End Enum

    Public Property Options2 As OptionsCustomized_e

End Class
