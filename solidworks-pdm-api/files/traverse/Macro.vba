Const VAULT_NAME As String = "MyVault"

Dim pdmVault As EdmVault5

Sub main()

    Set pdmVault = New EdmVault5
    pdmVault.LoginAuto VAULT_NAME, 0
    
    If pdmVault.IsLoggedIn Then
        
        Dim pdmFolder As IEdmFolder5
        
        Set pdmFolder = pdmVault.BrowseForFolder(0, "Select folder to traverse")
        
        If Not pdmFolder Is Nothing Then
            TraverseFolder pdmFolder
        End If
        
    Else
        Err.Raise vbError, "User is not logged in to the vault"
    End If
    
End Sub

Sub TraverseFolder(folder As IEdmFolder5, Optional parentLevel As String = "")

    Debug.Print parentLevel & "[+]" & folder.Name & " (" & folder.ID & ")"
    
    Dim thisLevel As String
    thisLevel = parentLevel & " "
    
    Dim pdmFilePos As IEdmPos5
    Set pdmFilePos = folder.GetFirstFilePosition()

    While Not pdmFilePos.IsNull
        Dim pdmFile As IEdmFile5
        Set pdmFile = folder.GetNextFile(pdmFilePos)
        Debug.Print thisLevel & " " & pdmFile.Name & " (" & pdmFile.ID & ")"
    Wend
    
    Dim pdmSubFolderPos As IEdmPos5
    Set pdmSubFolderPos = folder.GetFirstSubFolderPosition()
    
    While Not pdmSubFolderPos.IsNull
        Dim pdmSubFolder As IEdmFolder5
        Set pdmSubFolder = folder.GetNextSubFolder(pdmSubFolderPos)
        TraverseFolder pdmSubFolder, thisLevel
    Wend

End Sub