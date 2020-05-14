Const ARG_FORMAT = "__SwMacroArgs__"

Const GHND As Integer = &H42

Private Declare PtrSafe Function RegisterClipboardFormat Lib "User32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As Long

Public Sub SetArgument(arg As String)
    
    On Error GoTo ErrorHandler
        
    Dim wFormat As LongPtr
    
    wFormat = RegisterClipboardFormat(ARG_FORMAT)
    
    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
        
    hGlobalMemory = GlobalAlloc(GHND, Len(arg))
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    lpGlobalMemory = lstrcpy(lpGlobalMemory, arg)

    If GlobalUnlock(hGlobalMemory) <> 0 Then
        RaiseError "Failed to unlock memory"
    End If

    If OpenClipboard(0&) = 0 Then
        RaiseError "Failed to open clipboard"
    End If

    SetClipboardData wFormat, hGlobalMemory
    
    GoTo Finally
    
ErrorHandler:
    MsgBox "Critical Error: " & err.Description

Finally:
    CloseClipboard
    
End Sub

Sub RaiseError(desc As String)
    
    Const SYS_ERR_OFFSET As Integer = 513
    
    err.Raise Number:=vbObjectError + SYS_ERR_OFFSET, _
              Description:=desc
End Sub