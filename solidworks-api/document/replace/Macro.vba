#If VBA7 Then
     Private Declare PtrSafe Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal path As String) As Boolean
#Else
     Private Declare Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal Path As String) As boolean
#End If

Const REPLACEMENT_SUFFIX As String = ""
Const REPLACEMENT_PREFIX As String = "_"
Const REPLACEMENT_FILE_NAME As String = ""

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim srcFilePath As String
    
        srcFilePath = swModel.GetPathName()
        
        Dim replFilePath As String
        
        replFilePath = GetReplacementFilePath(srcFilePath)
    
        If Dir(replFilePath, vbNormal + vbReadOnly + vbHidden) <> "" Then
            
            If TypeOf swModel Is SldWorks.DrawingDoc Then
                Err.Raise vbError, "", "Not supported"
            Else
                Const S_OK As Integer = 1
                
                If swModel.ForceReleaseLocks() = S_OK Then
                    
                    FileCopy replFilePath, srcFilePath
                    
                    Dim res As Integer
                    res = swModel.ReloadOrReplace(False, "", True)
                    
                    If res <> swComponentReloadError_e.swReloadOkay Then
                        Err.Raise vbError, "", "Failed to reload model: " & res
                    End If
                    
                Else
                    Err.Raise vbError, "", "Failed to release file"
                End If
                
            End If
            
        Else
            Err.Raise vbError, "", "Replacement file path does not exist"
        End If
        
    Else
        Err.Raise vbError, "", "Open model"
    End If
    
End Sub

Function GetReplacementFilePath(srcFilePath As String) As String
    
    If srcFilePath <> "" Then
    
        Dim fileName As String
        
        If Not TryGetReplacementFileNameFromArguments(fileName) Then
            fileName = REPLACEMENT_FILE_NAME
        End If
        
        If fileName = "" Then
            fileName = REPLACEMENT_PREFIX & GetFileNameWithoutExtension(srcFilePath) & REPLACEMENT_SUFFIX & "." & GetExtension(srcFilePath)
        End If
        
        Dim filePath As String
        
        If PathIsRelative(fileName) Then
            filePath = GetFolderName(srcFilePath) & "\" & fileName
        Else
            filePath = fileName
        End If
        
        If LCase(filePath) <> LCase(srcFilePath) Then
            GetReplacementFilePath = filePath
        Else
            Err.Raise vbError, "", "Replacement file path and source file path match"
        End If
        
    Else
        Err.Raise vbError, "", "Model is not saved on disc"
    End If
    
End Function

Function TryGetReplacementFileNameFromArguments(ByRef fileName As String) As Boolean

try_:

    On Error GoTo catch_

    Dim macroOprMgr As Object
    Set macroOprMgr = CreateObject("CadPlus.MacroOperationManager")
        
    Set macroOper = macroOprMgr.PopOperation(swApp)
    
    Dim vArgs As Variant
    vArgs = macroOper.Arguments
   
    Dim macroArg As Object
    Set macroArg = vArgs(0)
    
    fileName = CStr(macroArg.GetValue())
    TryGetReplacementFileNameFromArguments = True
    GoTo finally_
    
catch_:
    TryGetReplacementFileNameFromArguments = False
finally_:

End Function

Function GetFolderName(filePath As String) As String
    GetFolderName = Left(filePath, InStrRev(filePath, "\") - 1)
End Function

Function GetFileNameWithoutExtension(filePath As String) As String
    GetFileNameWithoutExtension = Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
End Function

Function GetExtension(path As String) As String
    GetExtension = Right(path, Len(path) - InStrRev(path, "."))
End Function