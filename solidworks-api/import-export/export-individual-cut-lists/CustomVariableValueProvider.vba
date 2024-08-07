Option Explicit

Implements IMacroCustomVariableValueProvider

Function IMacroCustomVariableValueProvider_Provide(ByVal varName As String, ByVal args As Variant, ByVal context As Variant) As Variant

    Dim swCutList As SldWorks.Feature
    Set swCutList = context
    
    Select Case varName
        Case "cutListPrp":
            Dim prpName As String
            prpName = CStr(args(0))
            Dim swCustPrpsMgr As SldWorks.CustomPropertyManager
            Set swCustPrpsMgr = swCutList.CustomPropertyManager
            Dim prpVal As String
            swCustPrpsMgr.Get5 prpName, False, "", prpVal, False
            IMacroCustomVariableValueProvider_Provide = prpVal
        Case Else
            Err.Raise vbError, "", "Not supported variable: " & varName
    End Select

End Function