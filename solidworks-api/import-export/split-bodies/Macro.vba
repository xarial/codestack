Const BODY_FILTER As Integer = swBodyType_e.swAllBodies
Const VISIBLE_BODIES_ONLY As Boolean = True

Dim swApp As SldWorks.SldWorks


Sub main()

    Set swApp = Application.SldWorks
       
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        If swModel.GetType() = swDocumentTypes_e.swDocPART Then
               
            Dim outFolder As String
            outFolder = BrowseForFolder()
            
            If outFolder <> "" Then
                SplitBodies swModel, outFolder
            End If
            
        Else
            Err.Raise vbError, "", "Active document is not a part"
        End If
    
    Else
        Err.Raise vbError, "", "Open part file"
    End If
    
End Sub

Sub SplitBodies(model As SldWorks.ModelDoc2, outFolder As String)

    Dim enable3DInterconnect As Boolean
    Dim stepExportAppearances As Boolean
    Dim stepAP As Integer
    
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swMultiCAD_Enable3DInterconnect, False
    enable3DInterconnect = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swMultiCAD_Enable3DInterconnect)
    
    stepExportAppearances = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swStepExportAppearances)
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swStepExportAppearances, True
    
    stepAP = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP)
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swStepAP, 214
                   
try_:
    On Error GoTo catch_

    Dim swPart As SldWorks.PartDoc
    
    Set swPart = model
    
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(BODY_FILTER, VISIBLE_BODIES_ONLY)
    
    Dim i As Integer
    Dim swBody As SldWorks.Body2
    
    Dim errFiles As Integer
    Dim succFiles As Integer
    
    For i = 0 To UBound(vBodies)
    
        Set swBody = vBodies(i)
        
        Dim tempOutFilePath As String
        tempOutFilePath = GetTempFilePath(".step")
        
        If TryExportBody(model, swBody, tempOutFilePath) Then
            
            Dim outFilePath As String
            outFilePath = outFolder & "\" & GetFileNameWithoutExtension(model.GetTitle()) & "_" & swBody.Name & ".sldprt"
            
            If TryImportFileAs(tempOutFilePath, outFilePath) Then
                succFiles = succFiles + 1
            Else
                errFiles = errFiles + 1
            End If
            
            Kill tempOutFilePath
        Else
            errFiles = errFiles + 1
        End If
        
    Next
    
    MsgBox "Processed: " & succFiles & " file(s). Failed: " & errFiles & " file(s)"
    
    GoTo finally_
    
catch_:
    
    MsgBox Err.Description, vbCritical

finally_:

    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swMultiCAD_Enable3DInterconnect, enable3DInterconnect
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swStepExportAppearances, stepExportAppearances
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swStepAP, stepAP

End Sub

Function TryImportFileAs(srcFilePath As String, outFilePath As String) As Boolean
    
    Dim swModel As SldWorks.ModelDoc2
    
try_:
    On Error GoTo catch_
    
    Dim errs As Long
    Dim warns As Long
    
    Set swModel = swApp.LoadFile4(srcFilePath, "", Nothing, errs)
    
    If Not swModel Is Nothing Then
    
        If False <> swModel.Extension.SaveAs2(outFilePath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, "", False, errs, warns) Then
            TryImportFileAs = True
        Else
            Err.Raise vbError, "", "Failed to save file to'" & filePath & "'. Error code: " & errs
        End If
        
    Else
        Err.Raise vbError, "", "Failed to import file. Error code: " & errs
    End If
    
    GoTo finally_
catch_:
    TryImportFileAs = False
finally_:
    If Not swModel Is Nothing Then
        swApp.CloseDoc swModel.GetTitle
    End If
    
End Function

Function TryExportBody(model As SldWorks.ModelDoc2, body As SldWorks.Body2, filePath As String) As Boolean

try_:
    On Error GoTo catch_
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager
    
    swSelMgr.SuspendSelectionList
    
    Dim swBodies(0) As SldWorks.Body2
    Set swBodies(0) = body
    
    If swSelMgr.AddSelectionListObjects(swBodies, Nothing) = 1 Then
                
        Dim errs As Long
        Dim warns As Long
        Dim dir As String
        
        dir = GetDirectory(filePath)
        
        CreateDirectories dir
        
        If False <> model.Extension.SaveAs2(filePath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, "", False, errs, warns) Then
            TryExportBody = True
        Else
            Err.Raise vbError, "", "Failed to export '" & body.Name & "' to '" & filePath & "'. Error code: " & errs
        End If
    Else
        Err.Raise vbError, "", "Failed to select " & body.Name
    End If

    GoTo finally_
catch_:
    TryExportBody = False
finally_:

    swSelMgr.ResumeSelectionList2 False
    
End Function

Function GetTempFilePath(ext As String) As String
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim tempDirPath As String
    tempDirPath = fso.GetSpecialFolder(2)
    
    Dim uniqueName As String
    uniqueName = CreateObject("Scriptlet.TypeLib").GUID
    
    uniqueName = Mid(uniqueName, InStr(uniqueName, "{") + 1, InStrRev(uniqueName, "}") - InStrRev(uniqueName, "{") - 1)
    
    GetTempFilePath = tempDirPath & "\" & uniqueName & ext
    
End Function

Function GetExtension(path As String) As String
    GetExtension = Right(path, Len(path) - InStrRev(path, "."))
End Function

Function GetFileNameWithoutExtension(filePath As String) As String
    GetFileNameWithoutExtension = Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
End Function

Function GetDirectory(path As String)
    GetDirectory = Left(path, InStrRev(path, "\"))
End Function

Sub CreateDirectories(path As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(path) Then
        Exit Sub
    End If

    CreateDirectories fso.GetParentFolderName(path)
    
    fso.CreateFolder path
    
End Sub

Function BrowseForFolder(Optional title As String = "Select Output Folder") As String
    
    Dim shellApp As Object
    
    Set shellApp = CreateObject("Shell.Application")
    
    Dim folder As Object
    Set folder = shellApp.BrowseForFolder(0, title, 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.path
    Else
        BrowseForFolder = ""
    End If
    
End Function