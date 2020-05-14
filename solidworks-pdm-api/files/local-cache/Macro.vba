Const FILE_PATH As String = "FULL PATH TO FILE"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    GetLocalCopyFromVault FILE_PATH, "VAULT NAME", "USER NAME", "PASSWORD"
    
    Dim swOpenDocSpec As SldWorks.DocumentSpecification
    Set swOpenDocSpec = swApp.GetOpenDocSpec(FILE_PATH)
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.OpenDoc7(swOpenDocSpec)

End Sub

Sub GetLocalCopyFromVault(path As String, vaultName As String, userName As String, password As String)
    
    Dim pdmVault As New EdmLib.EdmVault5
    pdmVault.Login userName, password, vaultName
    
    If pdmVault.IsLoggedIn Then
        GetLocalCopies pdmVault, Array(path)
    Else
        MsgBox "Failed to login to vault"
    End If
    
End Sub

Sub GetLocalCopies(vault As EdmLib.EdmVault5, vFilePaths As Variant)
    
    If Not IsEmpty(vFilePaths) Then
        
        Dim pdmBatchGetUtil As EdmLib.IEdmBatchGet
        Set pdmBatchGetUtil = vault.CreateUtility(EdmUtil_BatchGet)
        
        Dim i As Integer
        
        Dim pdmSelItems() As EdmLib.EdmSelItem
        ReDim pdmSelItems(UBound(vFilePaths))
        
        For i = 0 To UBound(vFilePaths)
            
            Dim filePath As String
            filePath = vFilePaths(i)
            
            Dim pdmFile As EdmLib.IEdmFile5
            Dim pdmFolder As EdmLib.IEdmFolder5
            
            Set pdmFile = vault.GetFileFromPath(filePath, pdmFolder)
            
            pdmSelItems(i).mlDocID = pdmFile.ID
            pdmSelItems(i).mlProjID = pdmFolder.ID
            
        Next
        
        pdmBatchGetUtil.AddSelection vault, pdmSelItems
        pdmBatchGetUtil.CreateTree 0, EdmLib.EdmGetCmdFlags.Egcf_RefreshFileListing
        pdmBatchGetUtil.GetFiles 0
        
    End If
    
End Sub