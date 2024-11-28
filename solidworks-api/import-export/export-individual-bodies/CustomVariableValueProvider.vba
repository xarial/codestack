Option Explicit

Implements IMacroCustomVariableValueProvider

Function IMacroCustomVariableValueProvider_Provide(ByVal varName As String, ByVal args As Variant, ByVal context As Variant) As Variant

    Dim swBody As SldWorks.Body2
    Set swBody = context

    Select Case varName
        Case "bodyName":
            IMacroCustomVariableValueProvider_Provide = swBody.Name
        Case Else
            Err.Raise vbError, "", "Not supported variable: " & varName
    End Select

End Function