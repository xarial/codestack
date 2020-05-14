Const VAULT_NAME As String = "TestVault"

Dim swApp As SldWorks.SldWorks
Dim swPdmVault As IEdmVault5

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Set swPdmVault = New EdmVault5
        swPdmVault.LoginAuto VAULT_NAME, 0
        
        If swPdmVault.IsLoggedIn Then
            CheckOutModel swModel, swPdmVault
        Else
            MsgBox "Please login to vault"
        End If
    
    Else
        MsgBox "Please open the model"
    End If
    
End Sub

Sub CheckOutModel(model As SldWorks.ModelDoc2, vault As IEdmVault5)

    Dim modelPath As String
    modelPath = model.GetPathName()
    
    Dim swPdmFile As IEdmFile5
    Set swPdmFile = vault.GetFileFromPath(modelPath)

    If Not swPdmFile Is Nothing Then
        
        On Error GoTo catch

        Dim res As Boolean
        
        Dim swPdmFolder As IEdmFolder5
        Set swPdmFolder = vault.GetFolderFromPath(Left(modelPath, InStrRev(modelPath, "\")))

try:
        model.ForceReleaseLocks
        swPdmFile.LockFile swPdmFolder.ID, 0
        res = True
        GoTo finally
catch:
        Debug.Print Err.Number & ": "; Err.Description
        res = False
        GoTo finally
    
finally:
        model.ReloadOrReplace Not res, modelPath, Not res

    Else
        Err.Raise vbError, "", "Specified model doesn't exist in the vault"
    End If
    
End Sub
