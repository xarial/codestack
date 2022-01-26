Enum SearchType_e
    ByText
    ByExpression
    EmptyText
    EmptyExpression
    All
End Enum

Const FILTER As String = ""
Const SEARCH_TYPE As Integer = SearchType_e.EmptyText
Const USE_REGULAR_EXPRESSION As Boolean = False

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swApp.ActiveDoc
    
    If Not swDraw Is Nothing Then
           
        DeleteNotes swDraw
        
    Else
        Err.Raise vbError, "", "Only drawings are supported"
    End If
    
End Sub

Sub DeleteNotes(draw As SldWorks.DrawingDoc)
    
    Dim currentSheetName As String
    currentSheetName = draw.GetCurrentSheet().GetName
    
    Dim vSheets As Variant
    vSheets = draw.GetViews
    
    Dim i As Integer
        
    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        draw.ActivateSheet vViews(0).Name
        draw.ClearSelection2 False
        
        Dim j As Integer
        
        For j = 0 To UBound(vViews)
                
            Dim swView As SldWorks.View
            Set swView = vViews(j)
            
            Dim vNotes As Variant
            vNotes = swView.GetNotes
            
            Dim k As Integer
            
            For k = 0 To UBound(vNotes)
                
                Dim swNote As SldWorks.note
                Set swNote = vNotes(k)
                
                If ShouldDeleteNote(swNote) Then

                    Dim swAnn  As SldWorks.Annotation
                    Set swAnn = swNote.GetAnnotation
                    
                    Debug.Print "Deleting " & swNote.GetText & " (" & swNote.PropertyLinkedText & ")"

                    swAnn.Select3 True, Nothing
                    
                End If
                
            Next
            
        Next
        
        If draw.SelectionManager.GetSelectedObjectCount2(-1) > 0 Then
            If False <> draw.Extension.DeleteSelection2(swDeleteSelectionOptions_e.swDelete_Absorbed) Then
                draw.SetSaveFlag
            Else
                Err.Raise vbError, "", "Failed to delete annotations"
            End If
        End If
        
    Next
    
    draw.ActivateSheet currentSheetName
    
End Sub

Function ShouldDeleteNote(note As SldWorks.note) As Boolean

    Dim value As String
    
    Select Case SEARCH_TYPE
        Case SearchType_e.All
            ShouldDeleteNote = True
            Exit Function
        Case SearchType_e.EmptyText
            ShouldDeleteNote = note.GetText() = ""
            Exit Function
        Case SearchType_e.EmptyExpression
            ShouldDeleteNote = note.PropertyLinkedText = ""
            Exit Function
        Case SearchType_e.ByText
            value = note.GetText()
        Case SearchType_e.ByExpression
            value = note.PropertyLinkedText
    End Select
        
    If USE_REGULAR_EXPRESSION Then
        Dim regEx As Object
        Set regEx = CreateObject("VBScript.RegExp")
        
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.Pattern = FILTER
        
        ShouldDeleteNote = regEx.Test(value)
    Else
        ShouldDeleteNote = (value = FILTER)
    End If
    
End Function