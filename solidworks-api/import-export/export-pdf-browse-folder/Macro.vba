Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.ModelDoc2
    
    Set swDraw = swApp.ActiveDoc
    
    If swDraw Is Nothing Then
        Err.Raise vbError, "", "Open drawing"
    End If
    
    If swDraw.GetType() = swDocumentTypes_e.swDocDRAWING Then
    
        Dim outFolder As String
        outFolder = BrowseForFolder()
        
        If Right(outFolder, 1) = "\" Then
            outFolder = Left(outFolder, Len(outFolder) - 1)
        End If
        
        If outFolder <> "" Then
        
            Dim outFileName As String
            outFileName = GetFileNameWithoutExtension(swDraw.GetPathName()) & ".pdf"
            
            Dim outFilePath As String
            outFilePath = outFolder & "\" & outFileName
            
            Dim errs As Long
            Dim warns As Long
            
            If False = swDraw.Extension.SaveAs(outFilePath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, errs, warns) Then
                Err.Raise vbError, "", "Failed to export PDF to " & outFile
            End If
            
            If False <> swDraw.GetSaveFlag() Then
                If False = swDraw.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, errs, warns) Then
                    Err.Raise vbError, "", "Failed to save drawing"
                End If
            End If
        
            swApp.CloseDoc swDraw.GetTitle
            
        End If
    Else
        Err.Raise vbError, "", "Active document is not a drawing"
    End If
    
End Sub

Function GetFileNameWithoutExtension(filePath As String) As String
    GetFileNameWithoutExtension = Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
End Function

Function BrowseForFolder(Optional title As String = "Select Folder") As String
    
    Dim shellApp As Object
    
    Set shellApp = CreateObject("Shell.Application")
    
    Dim folder As Object
    Set folder = shellApp.BrowseForFolder(0, title, 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    End If
    
End Function