#If VBA7 Then
     Private Declare PtrSafe Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal path As String) As Boolean
#Else
     Private Declare Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal Path As String) As boolean
#End If

Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Boolean

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

Const DRAWINGS_FOLDER As String = ""

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
try_:
    On Error GoTo catch_
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        If swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY Then
            Dim swAssy As SldWorks.AssemblyDoc
            
            Set swAssy = swModel
            
            Dim vComps As Variant
            vComps = GetSelectedComponents(swModel.SelectionManager)
            
            If Not IsEmpty(vComps) Then
            
                Dim i As Integer
                Dim path As String
                path = vComps(0).GetPathName()
                
                For i = 1 To UBound(vComps)
                    If LCase(vComps(i).GetPathName()) <> LCase(path) Then
                        Err.Raise vbError, "", "Only identical components are supported"
                    End If
                Next
                
                Dim ext As String
                ext = Right(path, Len(path) - InStrRev(path, ".") + 1)
                
                Dim filter As String
                Dim fileType As String
                
                If LCase(ext) = ".sldprt" Then
                    fileType = "SOLIDWORKS Parts"
                ElseIf LCase(ext) = ".sldasm" Then
                    fileType = "SOLIDWORKS Assemblies"
                Else
                    Err.Raise vbError, "", "Unknown error"
                End If
                
                filter = fileType & " (*" & ext & ")|*" & ext & "|All Files (*.*)|*.*"
                
                Dim replaceFilePath As String
                replaceFilePath = BrowseForFileSave("Select replacement file path", filter, path)
                
                If replaceFilePath <> "" Then
                    If False = swAssy.MakeIndependent(replaceFilePath) Then
                        Err.Raise vbError, "", "Failed to make components independent"
                    End If
                    
                    MakeDrawingIndependent path, replaceFilePath
                    
                End If
            Else
                Err.Raise vbError, "", "Select components"
            End If
            
        Else
            Err.Raise vbError, "", "Only assembly documents are supported"
        End If
        
    Else
        Err.Raise vbError, "", "No model found"
    End If
    
    GoTo finally_
    
catch_:
    MsgBox Err.Description, vbCritical
finally_:
    
End Sub

Sub MakeDrawingIndependent(srcFilePath As String, destFilePath As String)
    
    Dim srcDrwFilePath As String
    srcDrwFilePath = ResolveDrawingPath(srcFilePath)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(srcDrwFilePath) Then
            
        Dim destDrwFilePath As String
        destDrwFilePath = ResolveDrawingPath(destFilePath)
            
        If fso.FileExists(destDrwFilePath) Then
            Err.Raise vbError, "", "Destination drawing already exists"
        End If
        
        fso.CopyFile srcDrwFilePath, destDrwFilePath, False
        
        Dim destDrwFilePathAttr As VbFileAttribute
        destDrwFilePathAttr = GetAttr(destDrwFilePath)
        
        If destDrwFilePathAttr And vbReadOnly Then
            Debug.Print "Removing the read-only flag from the destination drawing: " & destDrwFilePath
            SetAttr destDrwFilePath, destDrwFilePathAttr Xor vbReadOnly
        End If
        
        If False = swApp.ReplaceReferencedDocument(destDrwFilePath, srcFilePath, destFilePath) Then
            Err.Raise vbError, "", "Failed to replace referenced drawing document"
        End If
                
    End If
    
End Sub

Function ResolveDrawingPath(origFilePath As String) As String
    
    Dim targFolder As String
    
    If DRAWINGS_FOLDER = "" Then
        targFolder = GetFolderName(origFilePath)
    Else
        If PathIsRelative(DRAWINGS_FOLDER) Then
            targFolder = GetFolderName(origFilePath) & "\" & DRAWINGS_FOLDER
        Else
            targFolder = DRAWINGS_FOLDER
        End If
    End If
    
    ResolveDrawingPath = targFolder & "\" & GetFileNameWithoutExtension(origFilePath) & ".slddrw"

End Function

Function GetSelectedComponents(selMgr As SldWorks.SelectionMgr) As Variant

    Dim isInit As Boolean
    isInit = False
    
    Dim swComps() As SldWorks.Component2

    Dim i As Integer
    
    For i = 1 To selMgr.GetSelectedObjectCount2(-1)
                
        Dim swComp As SldWorks.Component2
    
        Set swComp = selMgr.GetSelectedObjectsComponent4(i, -1)
        
        If Not swComp Is Nothing Then
            
            If Not isInit Then
                ReDim swComps(0)
                Set swComps(0) = swComp
                isInit = True
            Else
                If Not Contains(swComps, swComp) Then
                    ReDim Preserve swComps(UBound(swComps) + 1)
                    Set swComps(UBound(swComps)) = swComp
                End If
            End If
                        
        End If
    
    Next

    If isInit Then
        GetSelectedComponents = swComps
    Else
        GetSelectedComponents = Empty
    End If

End Function

Function BrowseForFileSave(title As String, filters As String, initFilePath As String) As String
    
    Dim ofn As OPENFILENAME
    Const FILE_PATH_BUFFER_SIZE As Integer = 260
    
    Dim initFileName As String
    initFileName = Right(initFilePath, Len(initFilePath) - InStrRev(initFilePath, "\"))
    
    ofn.lpstrFilter = Replace(filters, "|", Chr(0)) & Chr(0)
    ofn.lpstrTitle = title
    ofn.nMaxFile = FILE_PATH_BUFFER_SIZE
    ofn.nMaxFileTitle = FILE_PATH_BUFFER_SIZE
    ofn.lpstrInitialDir = Left(initFilePath, InStrRev(initFilePath, "\") - 1)
    ofn.lpstrFile = initFileName & String(FILE_PATH_BUFFER_SIZE - Len(initFileName), Chr(0))
    ofn.lStructSize = LenB(ofn)
    
    Dim res As Boolean
    
    res = GetSaveFileName(ofn)

    If res Then
        
        Dim filePath As String
        filePath = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
        
        Dim vFilters As Variant
        vFilters = Split(filters, "|")
        Dim ext As String
        ext = vFilters((ofn.nFilterIndex - 1) * 2 + 1)
        ext = Right(ext, Len(ext) - InStrRev(ext, ".") + 1)
        
        If LCase(Right(filePath, Len(ext))) <> LCase(ext) Then
            filePath = filePath & ext
        End If
        
        BrowseForFileSave = filePath
        
    Else
        BrowseForFileSave = ""
    End If
    
End Function

Function Contains(vArr As Variant, item As Object) As Boolean
    
    Dim i As Integer
    
    For i = 0 To UBound(vArr)
        If vArr(i) Is item Then
            Contains = True
            Exit Function
        End If
    Next
    
    Contains = False
    
End Function

Function GetFolderName(filePath As String) As String
    GetFolderName = Left(filePath, InStrRev(filePath, "\") - 1)
End Function

Function GetFileNameWithoutExtension(filePath As String) As String
    GetFileNameWithoutExtension = Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
End Function