#Const ARGS = False 'True to use arguments from Toolbar+ or Batch+ instead of the constant

Const DATE_PRP_NAME As String = "Date"

Sub main()

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
        
try_:
    On Error GoTo catch_
    
    If Not swModel Is Nothing Then
        
        Dim dateFormat As String
        
        #If ARGS Then
                
            Dim macroRunner As Object
            Set macroRunner = CreateObject("CadPlus.MacroRunner.Sw")
            
            Dim param As Object
            Set param = macroRunner.PopParameter(swApp)
            
            Dim vArgs As Variant
            vArgs = param.Get("Args")
            
            dateFormat = CStr(vArgs(0))
        
        #Else
            dateFormat = GetDateFormat()
        #End If
    
        If dateFormat <> "" Then
            SetDateCustomProperty swModel, dateFormat
        End If
    Else
        Err.Raise vbError, "", "Please open model"
    End If
    
    GoTo finally_
catch_:
    MsgBox Err.Description, vbCritical
finally_:

End Sub

Function GetDateFormat(Optional defaultDateFormat As String = "dd/mm/yyyy") As String
    GetDateFormat = InputBox("Specify the format for the Date custom property", "Date Custom Property", defaultDateFormat)
End Function

Sub SetDateCustomProperty(model As SldWorks.ModelDoc2, dateFormat As String)
    
    Dim dateVal As String
    dateVal = Format(Now, dateFormat)
    
    Dim swCustPrpMgr As SldWorks.CustomPropertyManager
    
    Set swCustPrpMgr = model.Extension.CustomPropertyManager(confName)
    
    If swCustPrpMgr.Add3(DATE_PRP_NAME, swCustomInfoType_e.swCustomInfoText, dateVal, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue) <> swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged Then
        Err.Raise vbError, "", "Failed to add date property"
    End If
    
End Sub