#If VBA7 Then
     Private Declare PtrSafe Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal path As String) As Boolean
#Else
     Private Declare Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal Path As String) As boolean
#End If
        
Const MACROS_PATH As String = "Macro1.swp, D:\Macro2.swp, D:\MacrosFolder, Macros\Assembly"

Const PATH_DELIMETER As String = ","
Const MACRO_EXT As String = "swp"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
              
    Dim swMacrosColl As Collection
    Set swMacrosColl = New Collection
    
    AddMacros swMacrosColl
    
    Set swMacrosColl = ResolvePaths(swMacrosColl)
    
    RunMacros swMacrosColl

End Sub

Function ResolvePaths(swMacrosColl As Collection) As Collection

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim resColl As Collection
    Set resColl = New Collection
    
    Dim i As Integer
    
    For i = 1 To swMacrosColl.Count
        
        Dim path As String
        path = swMacrosColl(i)
        
        If PathIsRelative(path) Then
            path = fso.BuildPath(swApp.GetCurrentMacroPathFolder(), path)
        End If
        
        If fso.FolderExists(path) Then
            
            swMacrosColl.Remove i
            
            For Each file In fso.GetFolder(path).Files
                If LCase(fso.GetExtensionName(file)) = LCase(MACRO_EXT) Then
                    AddMacroToCollection resColl, file.path
                End If
            Next
            
        ElseIf fso.FileExists(path) Then
            AddMacroToCollection resColl, path
        Else
            Err.Raise vbObjectError, , "Macro file is not found: " & path
        End If
        
    Next
    
    Set ResolvePaths = resColl
    
End Function

Sub AddMacroToCollection(coll As Collection, item As String)
    
    If UCase(item) <> UCase(swApp.GetCurrentMacroPathName()) Then
        Dim i As Integer
        
        For i = 1 To coll.Count
            If UCase(coll.item(i)) = UCase(item) Then
                Exit Sub
            End If
        Next
        
        coll.Add item
    End If
    
End Sub

Sub RunMacros(swMacrosColl As Collection)
    
    Dim i As Integer
    
    For i = 1 To swMacrosColl.Count
        Dim path As String
        path = swMacrosColl(i)
        Dim macroErr As Long
        
        Dim moduleName As String
        Dim procName As String
        
        GetMacroEntryPoint path, moduleName, procName
        
        If False = swApp.RunMacro2(path, moduleName, procName, swRunMacroOption_e.swRunMacroUnloadAfterRun, macroErr) Then
            Err.Raise vbObjectError, , "Failed to run macro: " & path & ", error: " & macroErr
        End If
        
    Next
    
End Sub

Sub GetMacroEntryPoint(macroPath As String, ByRef moduleName As String, ByRef procName As String)
        
    Dim vMethods As Variant
    vMethods = swApp.GetMacroMethods(macroPath, swMacroMethods_e.swMethodsWithoutArguments)
    
    Dim i As Integer
    
    If Not IsEmpty(vMethods) Then
    
        For i = 0 To UBound(vMethods)
            Dim vData As Variant
            vData = Split(vMethods(i), ".")
            
            If i = 0 Or LCase(vData(1)) = "main" Then
                moduleName = vData(0)
                procName = vData(1)
            End If
        Next
        
    End If
    
End Sub

Sub AddMacros(swMacrosColl As Collection)
    
    Dim vPaths As Variant
    vPaths = Split(MACROS_PATH, PATH_DELIMETER)
    
    Dim i As Integer
    
    For i = 0 To UBound(vPaths)
    
        Dim path As String
        path = Trim(vPaths(i))
        swMacrosColl.Add path
        
    Next
    
End Sub