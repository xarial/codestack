Dim lblParamName() As Label
Dim txtParamValue() As TextBox

Dim WithEvents btnApply As CommandButton

Private Sub UserForm_Initialize()

    Me.Caption = "Edit " & FeatureName
    
    If UBound(DimensionNames) <> UBound(DimensionTitles) Then
        Err.Raise vbError, "", "Parameter names and dimensions must be of the same size"
    End If

    Dim i As Integer
    
    Dim maxWidth As Integer
    
    ReDim lblParamName(UBound(DimensionTitles))
    ReDim txtParamValue(UBound(DimensionTitles))
    
    Dim nextPosY As Integer
    nextPosY = MARGIN
    
    For i = 0 To UBound(DimensionTitles)
    
        Set lblParamName(i) = Me.Controls.Add("Forms.Label.1")
        lblParamName(i).Caption = CStr(DimensionTitles(i)) & ":"
        lblParamName(i).name = "lblLabel" & (i + 1)
        lblParamName(i).AutoSize = True
        
        lblParamName(i).Left = MARGIN
        lblParamName(i).Top = nextPosY
        
        If lblParamName(i).Width > maxWidth Then
            maxWidth = lblParamName(i).Width
        End If
        
        Set txtParamValue(i) = Me.Controls.Add("Forms.TextBox.1")
        txtParamValue(i).Width = TEXT_BOX_WIDTH
        txtParamValue(i).name = "txtVal" & (i + 1)
        txtParamValue(i).Top = nextPosY
                
        nextPosY = nextPosY + MARGIN + lblParamName(i).height
        
    Next
    
    For i = 0 To UBound(txtParamValue)
        txtParamValue(i).Left = maxWidth + MARGIN * 2
    Next
    
    Set btnApply = Me.Controls.Add("Forms.CommandButton.1")
    btnApply.Caption = "Apply"
    btnApply.name = "btnApply"
    btnApply.Top = nextPosY + MARGIN
    btnApply.Left = (maxWidth + MARGIN + TEXT_BOX_WIDTH) / 2 - btnApply.Width / 2 + MARGIN
    
    Dim height As Integer
    height = btnApply.Top + btnApply.height + MARGIN
    
    Me.StartUpPosition = 1 'center owner
    Me.ScrollBars = IIf(height > MAX_FORM_HEIGHT, fmScrollBarsVertical, fmScrollBarsNone)
    Me.ScrollHeight = height
    Me.Width = (maxWidth + MARGIN + TEXT_BOX_WIDTH) + MARGIN * 2 + 20
    Me.height = IIf(height > MAX_FORM_HEIGHT, MAX_FORM_HEIGHT + 25, height + 25) 'including header height
    
    LoadDimensionValues
    
End Sub

Private Sub LoadDimensionValues()
    
    Dim i As Integer
        
    For i = 0 To UBound(DimensionNames)
        
        Dim swDim As SldWorks.Dimension
        
        Dim dimName As String
        dimName = CStr(DimensionNames(i))
        
        Set swDim = GetDimension(dimName)
        
        If Not swDim Is Nothing Then
            Dim dimVal As Double
            Dim confNames(0) As String
            confNames(0) = ConfigName
            dimVal = swDim.GetValue3(swInConfigurationOpts_e.swSpecifyConfiguration, confNames)(0)
            txtParamValue(i).Text = dimVal
        Else
            Err.Raise vbError, "", dimName & " does not exist"
        End If
    Next
    
End Sub

Private Sub btnApply_Click()
    
    Dim i As Integer
        
    For i = 0 To UBound(DimensionNames)
        
        Dim swDim As SldWorks.Dimension
        
        Dim dimName As String
        dimName = CStr(DimensionNames(i))
        
        Set swDim = GetDimension(dimName)
        
        If Not swDim Is Nothing Then
            Dim dimVal As Double
            
            If IsNumeric(txtParamValue(i).Text) Then
                dimVal = CDbl(txtParamValue(i).Text)
            Else
                Err.Raise vbError, "", "Specified value for " & DimensionTitles(i) & " is not numeric"
            End If
            Dim confNames(0) As String
            confNames(0) = ConfigName
            swDim.SetValue3 dimVal, swInConfigurationOpts_e.swSpecifyConfiguration, confNames
        Else
            Err.Raise vbError, "", dimName & " does not exist"
        End If
    Next
    
    ActiveModel.ForceRebuild3 False
    
End Sub

Function GetDimension(name As String) As SldWorks.Dimension
    
    Dim dimParts As Variant
    dimParts = Split(name, "/")
    
    Dim i As Integer
    
    Dim swTargetModel As SldWorks.ModelDoc2
    Set swTargetModel = Model
    
    Dim swCurComp As SldWorks.Component2
    
    For i = 0 To UBound(dimParts) - 1
        Dim swAssy As SldWorks.AssemblyDoc
        Set swAssy = swTargetModel
        Set swCurComp = swAssy.GetComponentByName(dimParts(i))
        Set swTargetModel = swCurComp.GetModelDoc2()
    Next
    
    Set GetDimension = swTargetModel.Parameter(dimParts(UBound(dimParts)))
    
End Function