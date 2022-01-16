Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    If Not swDraw Is Nothing Then
        
        Dim swSheet As SldWorks.sheet
        Set swSheet = swDraw.GetCurrentSheet
        
        Dim swDefView As SldWorks.View
        Set swDefView = GetDefaultView(swSheet)
        
        If Not swDefView Is Nothing Then
            
            Dim swRefDoc As SldWorks.ModelDoc2
            Set swRefDoc = swDefView.ReferencedDocument
            
            If Not swRefDoc Is Nothing Then
            
                ValidateSheet swSheet, swRefDoc
                
                Dim vConfNames As Variant
                vConfNames = swRefDoc.GetConfigurationNames
                
                Dim i As Integer
                
                For i = 0 To UBound(vConfNames)
                    
                    Dim confName As String
                    confName = CStr(vConfNames(i))
                    
                    If LCase(swDefView.ReferencedConfiguration) <> LCase(confName) Then
                        CopySheet swDraw, swSheet, confName
                    End If
                    
                Next
                
            Else
                Err.Raise vbError, "", "Default view does not have referenced document"
            End If
            
        Else
            Err.Raise vbError, "", "Default view is not found"
        End If
        
    Else
        Err.Raise vbError, "", "Open drawing"
    End If
    
End Sub

Function GetDefaultView(sheet As SldWorks.sheet) As SldWorks.View
    
    Dim vViews As Variant
    
    vViews = sheet.GetViews
    
    If Not IsEmpty(vViews) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(vViews)
            
            Dim swView As SldWorks.View
            Set swView = vViews(i)
            
            If UCase(swView.Name) = UCase(sheet.CustomPropertyView) Then
                Set GetDefaultView = swView
                Exit Function
            End If
            
        Next
        
        Set GetDefaultView = vViews(0) 'use first one
    Else
        Set GetDefaultView = Nothing
    End If
    
End Function

Sub ValidateSheet(sheet As SldWorks.sheet, refDoc As SldWorks.ModelDoc2)
    
    Dim vViews As Variant
    vViews = sheet.GetViews
    
    Dim i As Integer
    
    For i = 0 To UBound(vViews)
        
        Dim swView As SldWorks.View
        Set swView = vViews(i)
        
        If Not swView.ReferencedDocument Is refDoc Then
            Err.Raise vbError, "", "Different models are referenced in " & sheet.GetName
        End If
        
    Next
    
End Sub

Sub CopySheet(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet, confName As String)
    
    If False <> draw.Extension.SelectByID2(sheet.GetName(), "SHEET", 0, 0, 0, False, 0, Nothing, 0) Then
        
        draw.EditCopy
        
        If False <> draw.PasteSheet(swInsertOptions_e.swInsertOption_MoveToEnd, swRenameOptions_e.swRenameOption_Yes) Then
            
            Dim swNewSheet As SldWorks.sheet
            Set swNewSheet = draw.sheet(draw.GetSheetNames()(draw.GetSheetCount() - 1))
            
            Dim vViews As Variant
            vViews = swNewSheet.GetViews
            
            Dim i As Integer
            
            For i = 0 To UBound(vViews)
                
                Dim swView As SldWorks.View
                Set swView = vViews(i)
                
                swView.ReferencedConfiguration = confName
                
            Next
            
            swNewSheet.SetName confName
            
        Else
            Err.Raise vbError, "", "Failed to paste sheet"
        End If
    Else
        Err.Raise vbError, "", "Failed to select sheet"
    End If

End Sub