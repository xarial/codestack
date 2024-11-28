#Const MACRO_PLUS = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2

    Dim macroOper As Object
    
#If MACRO_PLUS Then
    
    Dim operMgr As Object
    Set operMgr = CreateObject("CadPlus.MacroOperationManager")
    
    Set macroOper = operMgr.PopOperation(swApp)
    
    Set swModel = macroOper.Model
    
#Else
    Set swModel = swApp.ActiveDoc
#End If

    Dim vConfNames As Variant
    
    vConfNames = swModel.GetConfigurationNames
    
    Dim hasSmConfs As Boolean
    Dim deletedConfsList As String
    
    If Not IsEmpty(vConfNames) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(vConfNames)
            
            Dim confName As String
            confName = CStr(vConfNames(i))
            
            Dim swConf As SldWorks.Configuration
            Set swConf = swModel.GetConfigurationByName(confName)
            
            If swConf.Type = swConfigurationType_e.swConfiguration_SheetMetal Then
                
                hasSmConfs = True
                
                If False <> swModel.DeleteConfiguration2(swConf.Name) Then
                    
                    If deletedConfsList <> "" Then
                        deletedConfsList = deletedConfsList & vbLf
                    End If
                    
                    deletedConfsList = deletedConfsList & swConf.Name
                    
                Else
                    #If MACRO_PLUS Then
                        macroOper.ReportIssue "Failed to delete configuration '" & confName & "'", 2
                        macroOper.SetStatus 4
                    #End If
                End If
                
            End If
            
        Next
        
    End If

#If MACRO_PLUS Then
    If hasSmConfs Then
        If deletedConfsList <> "" Then
            macroOper.SetResult deletedConfsList
        Else
            macroOper.SetStatus 2
        End If
    Else
        macroOper.ReportIssue "No sheet metal configurations found", 1
        macroOper.SetStatus 4
    End If
#End If

End Sub