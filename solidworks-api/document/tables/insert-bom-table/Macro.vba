Const ANCHOR_TYPE As Integer = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft
Const BOM_TYPE As Integer = swBomType_e.swBomType_PartsOnly
Const TABLE_TEMPLATE As String = ""
Const INDENTED_NUMBERING_TYPE As Integer = swNumberingType_e.swNumberingType_Flat
Const DETAILED_CUT_LIST As Boolean = False
Const FOLLOW_ASSEMBLY_ORDER As Boolean = True

Const ALL_SHEETS As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    If ALL_SHEETS Then
    
        Dim vSheetNames As Variant
        vSheetNames = swDraw.GetSheetNames
        
        Dim activeSheetName As String
        activeSheetName = swDraw.GetCurrentSheet().GetName
        
        Dim i As Integer
        
        For i = 0 To UBound(vSheetNames)
            Dim swSheet As SldWorks.sheet
            Set swSheet = swDraw.sheet(CStr(vSheetNames(i)))
            InsertBomTable swDraw, swSheet
        Next
        
        swDraw.ActivateSheet activeSheetName
        
    Else
        InsertBomTable swDraw, swDraw.GetCurrentSheet
    End If
    
End Sub

Sub InsertBomTable(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet)
    
    If False = draw.ActivateSheet(sheet.GetName()) Then
        Err.Raise vbError, "", "Failed to activate sheet " & sheet.GetName
    End If
    
    Dim vViews As Variant
    vViews = sheet.GetViews
    
    Dim swView As SldWorks.View
    
    Set swView = vViews(0)
    
    Dim swBomTableAnn As SldWorks.BomTableAnnotation
    
    Set swBomTableAnn = swView.InsertBomTable4(True, 0, 0, ANCHOR_TYPE, BOM_TYPE, "", TABLE_TEMPLATE, False, INDENTED_NUMBERING_TYPE, DETAILED_CUT_LIST)
        
    If Not swBomTableAnn Is Nothing Then
        swBomTableAnn.BomFeature.FollowAssemblyOrder2 = FOLLOW_ASSEMBLY_ORDER
    Else
        Err.Raise vbError, "", "Failed to insert BOM table into " & swView.Name
    End If
    
End Sub