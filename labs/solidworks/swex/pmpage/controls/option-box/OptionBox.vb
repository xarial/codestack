Imports CodeStack.SwEx.Common.Attributes
Imports CodeStack.SwEx.PMPage.Attributes

Public Class OptionBoxDataModel

    Public Enum Options_e
        Option1
        Option2
        <Title("Third Option")>
        Option3
    End Enum

    <OptionBox>
    Public Property Options As Options_e

End Class
