Const TOP_LEVEL_CONFIGS_ONLY As Boolean = False
Const USE_CORRESPONDING_FLAT_PATTERN_CONF As Boolean = True
Const GENERATE_MISSING_FLAT_PATTERN_CONF As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    If Not swDraw Is Nothing Then
        
        Dim swSheet As SldWorks.sheet
        Set swSheet = swDraw.GetCurrentSheet
        
        Dim swDefView As SldWorks.view
        Set swDefView = GetDefaultView(swDraw, swSheet)
        
        If Not swDefView Is Nothing Then
            
            Dim swRefDoc As SldWorks.ModelDoc2
            Set swRefDoc = swDefView.ReferencedDocument
            
            If Not swRefDoc Is Nothing Then
            
                ValidateSheet swSheet, swRefDoc
                
                Dim vConfNames As Variant
                vConfNames = GetConfigurations(swRefDoc)
                
                Dim i As Integer
                
                For i = 0 To UBound(vConfNames)
                    
                    Dim confName As String
                    confName = CStr(vConfNames(i))
                    
                    If LCase(GetActualReferencedConfiguration(swDefView)) <> LCase(confName) Then
                        CopySheetWithConfiguration swDraw, swSheet, confName
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

Function GetConfigurations(refDoc As SldWorks.ModelDoc2) As Variant
    
    Dim confNames() As String
    
    Dim vConfNames As Variant
    vConfNames = refDoc.GetConfigurationNames
    
    Dim i As Integer
    
    For i = 0 To UBound(vConfNames)
        
        Dim confName As String
        confName = CStr(vConfNames(i))
        
        Dim swConf As SldWorks.Configuration
        Set swConf = refDoc.GetConfigurationByName(confName)
        
        If (Not TOP_LEVEL_CONFIGS_ONLY Or swConf.GetParent() Is Nothing) And swConf.Type = swConfigurationType_e.swConfiguration_Standard Then
                
            If (Not confNames) = -1 Then
                ReDim confNames(0)
            Else
                ReDim Preserve confNames(UBound(confNames) + 1)
            End If
            
            confNames(UBound(confNames)) = confName
            
        End If
        
    Next
    
    GetConfigurations = confNames
    
End Function

Function GetActualReferencedConfiguration(view As SldWorks.view) As String
    
    Dim refConfName As String
    refConfName = view.ReferencedConfiguration
    
    Dim swConf As SldWorks.Configuration
    
    Set swConf = view.ReferencedDocument.GetConfigurationByName(refConfName)
    
    If swConf.Type <> swConfigurationType_e.swConfiguration_Standard Then
        Set swConf = swConf.GetParent
    End If
    
    GetActualReferencedConfiguration = swConf.Name
    
End Function

Function GetDefaultView(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet) As SldWorks.view
    
    Dim vViews As Variant
    
    vViews = GetSheetViews(draw, sheet)
    
    If Not IsEmpty(vViews) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(vViews)
            
            Dim swView As SldWorks.view
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
        
        Dim swView As SldWorks.view
        Set swView = vViews(i)
        
        If Not swView.ReferencedDocument Is refDoc Then
            Err.Raise vbError, "", "Different models are referenced in " & sheet.GetName
        End If
        
    Next
    
End Sub

Sub CopySheetWithConfiguration(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet, baseConfName As String)
    
    Const MAX_PASTE_ATEMPTS As Integer = 3
    
    If False <> draw.Extension.SelectByID2(sheet.GetName(), "SHEET", 0, 0, 0, False, 0, Nothing, 0) Then
        
        draw.EditCopy
        
        If TryPasteSheet(draw, MAX_PASTE_ATEMPTS) Then
            
            Dim swNewSheet As SldWorks.sheet
            Set swNewSheet = draw.sheet(draw.GetSheetNames()(draw.GetSheetCount() - 1))
            
            Dim vViews As Variant
            vViews = GetSheetViews(draw, swNewSheet)
            
            Dim i As Integer
            
            For i = 0 To UBound(vViews)
                
                Dim swView As SldWorks.view
                Set swView = vViews(i)
                
                Dim confName As String
                
                If False <> swView.IsFlatPatternView() And USE_CORRESPONDING_FLAT_PATTERN_CONF Then
                    confName = GetFlatPatternConfiguration(draw, swView.ReferencedDocument, baseConfName, GENERATE_MISSING_FLAT_PATTERN_CONF)
                Else
                    confName = baseConfName
                End If
                
                swView.ReferencedConfiguration = confName
                
                RefreshView draw, swView
                
            Next
            
            swNewSheet.SetName baseConfName
                        
        Else
            Err.Raise vbError, "", "Failed to paste sheet"
        End If
    Else
        Err.Raise vbError, "", "Failed to select sheet"
    End If

End Sub

Function TryPasteSheet(draw As SldWorks.DrawingDoc, attempts As Integer) As Boolean

    Dim curAttemp As Integer
    curAttemp = 1

    'It was observed than in some cases first atempt to paste sheet fails
    While False = draw.PasteSheet(swInsertOptions_e.swInsertOption_MoveToEnd, swRenameOptions_e.swRenameOption_Yes)
        
        Debug.Print "Failed to paste a sheet on atttempt: " & curAttemp
        
        If curAttemp >= attempts Then
            TryPasteSheet = False
            Exit Function
        End If
        
        curAttemp = curAttemp + 1
 
    Wend
    
    TryPasteSheet = True

End Function

'In some cases new configuration of view is not updated until refreshed
Sub RefreshView(draw As SldWorks.DrawingDoc, swView As SldWorks.view)
    
    If SelectDrawingView(draw, swView) Then
        
        draw.SuppressView
        
        If SelectDrawingView(draw, swView) Then
            draw.UnsuppressView
        End If
        
    End If
    
End Sub

Function GetFlatPatternConfiguration(draw As SldWorks.DrawingDoc, refDoc As SldWorks.ModelDoc2, baseConfName As String, allowCreateIfNotExist As Boolean) As String
    
    Dim swConf As SldWorks.Configuration
    Set swConf = refDoc.GetConfigurationByName(baseConfName)
    
    If swConf.Type <> swConfigurationType_e.swConfiguration_SheetMetal Then
        
        Dim vChildrenConfs As Variant
        
        vChildrenConfs = swConf.GetChildren()
        
        Dim i As Integer
        
        If Not IsEmpty(vChildrenConfs) Then
        
            For i = 0 To UBound(vChildrenConfs)
                
                Dim swChildConf As SldWorks.Configuration
                Set swChildConf = vChildrenConfs(i)
                
                If swChildConf.Type = swConfigurationType_e.swConfiguration_SheetMetal Then
                    Debug.Print "Using flat pattern configuration " & swChildConf.Name & " for the " & baseConfName
                    GetFlatPatternConfiguration = swChildConf.Name
                    Exit Function
                End If
                
            Next
        
        End If
        
        If allowCreateIfNotExist Then
            Debug.Print "Creating flat pattern configuration for " & baseConfName
            GetFlatPatternConfiguration = CreateFlatPatternConfiguration(draw, refDoc, baseConfName)
        Else
            Debug.Print "Flat pattern configuration is not found for " & baseConfName
            GetFlatPatternConfiguration = baseConfName
        End If
    Else
        GetFlatPatternConfiguration = baseConfName
    End If
    
End Function

Function CreateFlatPatternConfiguration(draw As SldWorks.DrawingDoc, refDoc As SldWorks.ModelDoc2, baseConfName As String) As String
    
    Dim swFlatPatternView As SldWorks.view
    Set swFlatPatternView = draw.CreateFlatPatternViewFromModelView3(refDoc.GetPathName(), baseConfName, 0, 0, 0, True, False)
    
    If Not swFlatPatternView Is Nothing Then
        CreateFlatPatternConfiguration = swFlatPatternView.ReferencedConfiguration
        
        If SelectDrawingView(draw, swFlatPatternView) Then
            If False = draw.Extension.DeleteSelection2(swDeleteSelectionOptions_e.swDelete_Absorbed) Then
                Err.Raise vbError, "", "Failed to delete temp view"
            End If
        Else
            Err.Raise vbError, "", "Failed to select temp view for deletion"
        End If
        
    Else
        Err.Raise vbError, "", "Failed to create temp flat pattern view for " & refDoc.GetPathName() & " (" & baseConfName & ")"
    End If
    
End Function

Function SelectDrawingView(draw As SldWorks.ModelDoc2, view As SldWorks.view) As Boolean
    SelectDrawingView = False <> draw.Extension.SelectByID2(view.Name, "DRAWINGVIEW", 0, 0, 0, False, -1, Nothing, swSelectOption_e.swSelectOptionDefault)
End Function

Function GetSheetViews(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet) As Variant
    
    'ISheet::GetViews also returns views from the view palette
    
    Dim vSheets As Variant
    
    vSheets = draw.GetViews
    
    Dim i As Integer
    
    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.view
        Set swSheetView = vViews(0)
        
        If swSheetView.GetName2() = sheet.GetName() Then
            
            If UBound(vViews) > 0 Then
                
                Dim swViews() As SldWorks.view
                ReDim swViews(UBound(vViews) - 1)
                
                Dim j As Integer
                
                For j = 0 To UBound(swViews)
                    Set swViews(j) = vViews(j + 1)
                Next
                
                GetSheetViews = swViews
                Exit Function
                
            Else
                Err.Raise vbError, "", "No drawing view found in " & sheet.GetName
            End If
            
        End If
            
    Next
    
    Err.Raise vbError, "", "Failed to get drawing views from " & sheet.GetName
    
End Function