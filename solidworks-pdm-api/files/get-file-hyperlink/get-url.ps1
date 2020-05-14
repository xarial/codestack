$vault_name=$args[0]
$filePath=$args[1]
$action=$args[2]

$Source = @"
Imports System
Imports System.IO

Public Class SwPdmTools

    'open view explore get lock properties history
    Public Shared Sub GetHyperlink(vaultName As String, filePath As String, action As String)

        Dim vault As Object = Activator.CreateInstance(Type.GetTypeFromProgID("ConisioLib.EdmVault"))

        vault.LoginAuto(vaultName, 0)

        Dim folderPath As String = Path.GetDirectoryName(filePath)
        Dim fileName As String = Path.GetFileName(filePath)

        Dim folder As Object = vault.GetFolderFromPath(folderPath)
        Dim file As Object = folder.GetFile(fileName)

        If Not file Is Nothing Then

            Const EdmObject_File As Integer = 1
            Dim url As String = String.Format("conisio://{0}/{1}?projectid={2}&documentid={3}&objecttype={4}", vaultName, action, folder.ID, file.ID, EdmObject_File)
            Console.WriteLine(url)

        End If

    End Sub

End Class
"@

Add-Type -TypeDefinition $Source -Language VisualBasic

[SwPdmTools]::GetHyperlink($vault_name, $filePath, $action)
