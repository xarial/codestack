Imports EPDM.Interop.epdm

Module Module1

    Sub Main()
        Try
            Dim vault As New EdmVault5

            Dim args As String() = Environment.GetCommandLineArgs()

            Dim vaultName As String = args(1)
            Dim login As String = args(2)
            Dim password As String = args(3)

            vault.Login(login, password, vaultName)

            If vault.IsLoggedIn Then
                CheckInAllCheckedOutFiles(vault)
            Else
                Throw New Exception("Failed to login to vault")
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Sub CheckInAllCheckedOutFiles(vault As EdmVault5)

        Dim checkOutFiles = GetCheckedOutFilesList(vault)

        If checkOutFiles.Any() Then

            Console.WriteLine(String.Format("Checking in {0} files", checkOutFiles.Length))

            Dim selItems(checkOutFiles.Length - 1) As EdmSelItem

            For i As Integer = 0 To checkOutFiles.Length - 1
                selItems(i) = New EdmSelItem()
                selItems(i).mlDocID = checkOutFiles(i)
            Next

            Dim batchUnlockUtil As IEdmBatchUnlock2 = vault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock)

            batchUnlockUtil.AddSelection(vault, selItems)

            batchUnlockUtil.CreateTree(IntPtr.Zero, EdmUnlockBuildTreeFlags.Eubtf_MayUnlock + EdmUnlockBuildTreeFlags.Eubtf_MayUndoLock + EdmUnlockBuildTreeFlags.Eubtf_UndoLockDefault + EdmUnlockBuildTreeFlags.Eubtf_RefreshFileListing)

            batchUnlockUtil.GetFileList(EdmUnlockFileListFlag.Euflf_GetUnlocked + EdmUnlockFileListFlag.Euflf_GetUndoLocked + EdmUnlockFileListFlag.Euflf_GetUnprocessed)

            batchUnlockUtil.UnlockFiles(IntPtr.Zero, Nothing)
        Else
            Console.WriteLine("There are not files to check-in")
        End If

    End Sub

    Function GetCheckedOutFilesList(vault As EdmVault5) As Integer()

        Dim fileIds As New List(Of Integer)
        Dim search As IEdmSearch6

        search = vault.CreateSearch()

        search.FindFiles = True
        search.FindFolders = False
        search.Recursive = True
        search.FindLockedFiles = True
        search.FindUnlockedFiles = False

        Dim searchRes As IEdmSearchResult5 = search.GetFirstResult

        While Not searchRes Is Nothing
            Console.WriteLine(String.Format("{0} [{1}] - {2}", searchRes.Path, searchRes.ID, searchRes.LockedByUserName))
            fileIds.Add(searchRes.ID)
            searchRes = search.GetNextResult()
        End While

        Return fileIds.ToArray()

    End Function

End Module