Const FILE_PATH As String = ""
Const PRINT_FILE_NAME As Boolean = True
Const PRINT_VIEW_NAME As Boolean = True
Const FILTER As String = ""

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swApp.ActiveDoc
    
    If Not swDraw Is Nothing Then
        
        Dim outFilePath As String
        
        If FILE_PATH <> "" Then
            outFilePath = FILE_PATH
        Else
            outFilePath = swDraw.GetPathName
            
            If outFilePath = "" Then
                Err.Raise "Drawing is not saved to the and FILE_PATH is not specified"
            End If
            
            outFilePath = Left(outFilePath, InStrRev(outFilePath, ".") - 1) & "_Notes.txt"
        End If
        
        Dim fileNmb As Integer
        fileNmb = FreeFile
    
        Open outFilePath For Append As #fileNmb
    
        If PRINT_FILE_NAME Then
            Print #fileNmb, "*** File Path: " & swDraw.GetPathName & " ***"
        End If
    
        PrintNotes swDraw, fileNmb
        
        Print #fileNmb, ""
        Close #fileNmb
        
    Else
        Err.Raise vbError, "", "Only drawings are supported"
    End If
    
End Sub

Sub PrintNotes(draw As SldWorks.DrawingDoc, fileNmb As Integer)
    
    Dim vSheets As Variant
    vSheets = draw.GetViews
    
    Dim i As Integer
        
    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim j As Integer
        
        For j = 0 To UBound(vViews)
            
            Dim swView As SldWorks.View
            Set swView = vViews(j)
            
            If PRINT_VIEW_NAME Then
                Print #fileNmb, "*** View Name: " & swView.Name & " ***"
            End If
            
            Dim vNotes As Variant
            vNotes = swView.GetNotes
            
            Dim k As Integer
            
            For k = 0 To UBound(vNotes)
                Dim swNote As SldWorks.Note
                Set swNote = vNotes(k)
                
                Dim text As String
                text = swNote.GetText
                
                If IncludeNote(text) Then
                    If text = "" Then
                        text = "[X]"
                    End If
                    
                    Print #fileNmb, text
                End If
                
            Next
            
        Next
        
    Next
    
End Sub

Function IncludeNote(text As String) As Boolean

    If FILTER = "" Then
        IncludeNote = True
    Else
        Dim regEx As Object
        Set regEx = CreateObject("VBScript.RegExp")
        
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.Pattern = FILTER
        
        IncludeNote = regEx.Test(text)
    
    End If
    
End Function