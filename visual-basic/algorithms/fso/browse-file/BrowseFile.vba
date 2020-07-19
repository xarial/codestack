Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Boolean

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As LongPtr
  hInstance As LongPtr
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As LongPtr
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Const FILTER As String = "Text Files (*.txt)|*.txt|PNG Image Files (*.png)|*.png|All Files (*.*)|*.*"

Sub main()

    Dim filePath As String
    filePath = BrowseForFileSave("Select file path to save", FILTER)
    
    If filePath <> "" Then
        Debug.Print "Selected save file path: " & filePath
    Else
        Debug.Print "No save file selected"
    End If
    
    filePath = BrowseForFileOpen("Select file path to open", FILTER)
    
    If filePath <> "" Then
        Debug.Print "Selected open file path: " & filePath
    Else
        Debug.Print "No open file selected"
    End If

End Sub

Function BrowseForFileSave(title As String, filters As String) As String
    BrowseForFileSave = BrowseForFile(title, filters, True)
End Function

Function BrowseForFileOpen(title As String, filters As String) As String
    BrowseForFileOpen = BrowseForFile(title, filters, False)
End Function

Function BrowseForFile(title As String, filters As String, save As Boolean) As String
    
    Dim ofn As OPENFILENAME
    Const FILE_PATH_BUFFER_SIZE As Integer = 260
    
    ofn.lpstrFilter = Replace(filters, "|", Chr(0)) & Chr(0)
    ofn.lpstrTitle = title
    ofn.nMaxFile = FILE_PATH_BUFFER_SIZE
    ofn.nMaxFileTitle = FILE_PATH_BUFFER_SIZE
    ofn.lpstrFile = String(FILE_PATH_BUFFER_SIZE, Chr(0))
    ofn.lStructSize = LenB(ofn)
    
    Dim res As Boolean
    
    If save Then
        res = GetSaveFileName(ofn)
    Else
        res = GetOpenFileName(ofn)
    End If
    
    If res Then
        
        Dim filePath As String
        filePath = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
        
        If save Then
            Dim vFilters As Variant
            vFilters = Split(FILTER, "|")
            Dim ext As String
            ext = vFilters((ofn.nFilterIndex - 1) * 2 + 1)
            ext = Right(ext, Len(ext) - InStrRev(ext, ".") + 1)
            
            If LCase(Right(filePath, Len(ext))) <> LCase(ext) Then
                filePath = filePath & ext
            End If
        End If
        
        BrowseForFile = filePath
        
    Else
        BrowseForFile = ""
    End If
    
End Function