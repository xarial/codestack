Dim lblParamName() As Label
Dim txtParamValue() As TextBox

Dim WithEvents chkCreateConf As CheckBox
Dim txtConfName As TextBox
Dim WithEvents btnApply As CommandButton

Dim FeatDimsInfos() As DimensionInfo
Dim swActiveModel As SldWorks.ModelDoc2
Dim FeatConfName As String

Public Sub EditDimensions(dimsInfos() As DimensionInfo, activeModel As SldWorks.ModelDoc2, confName As String)
    
    LoadLayout dimsInfos, activeModel, confName
    
    Me.Show vbModeless
    
End Sub

Private Sub LoadLayout(dimsInfos() As DimensionInfo, activeModel As SldWorks.ModelDoc2, confName As String)
    
    FeatDimsInfos = dimsInfos
    FeatConfName = confName
    
    Set swActiveModel = activeModel
    
    Dim i As Integer
    
    Dim maxWidth As Integer
    
    ReDim lblParamName(UBound(FeatDimsInfos))
    ReDim txtParamValue(UBound(FeatDimsInfos))
    
    Dim nextPosY As Integer
    nextPosY = MARGIN
    
    For i = 0 To UBound(FeatDimsInfos)
        
        Dim dimInfo As DimensionInfo
        dimInfo = FeatDimsInfos(i)
        
        Set lblParamName(i) = Me.Controls.Add("Forms.Label.1")
        lblParamName(i).Caption = dimInfo.title & ":"
        lblParamName(i).Name = "lblLabel" & (i + 1)
        lblParamName(i).AutoSize = True
        
        lblParamName(i).Left = MARGIN
        lblParamName(i).Top = nextPosY
        
        If lblParamName(i).Width > maxWidth Then
            maxWidth = lblParamName(i).Width
        End If
        
        Set txtParamValue(i) = Me.Controls.Add("Forms.TextBox.1")
        txtParamValue(i).Width = TEXT_BOX_WIDTH
        txtParamValue(i).Name = "txtVal" & (i + 1)
        txtParamValue(i).Top = nextPosY
        txtParamValue(i).Text = dimInfo.Value
        
        nextPosY = nextPosY + MARGIN + lblParamName(i).height
        
    Next
    
    For i = 0 To UBound(txtParamValue)
        txtParamValue(i).Left = maxWidth + MARGIN * 2
    Next
    
    Set chkCreateConf = Me.Controls.Add("Forms.CheckBox.1")
    chkCreateConf.Caption = "Create Configuration"
    chkCreateConf.Name = "chkCreateConf"
    chkCreateConf.Top = nextPosY + MARGIN
    chkCreateConf.Left = MARGIN
    
    Set txtConfName = Me.Controls.Add("Forms.TextBox.1")
    txtConfName.Name = "txtConfName"
    txtConfName.Top = chkCreateConf.Top + chkCreateConf.height + MARGIN
    txtConfName.Left = MARGIN
    txtConfName.Text = FeatConfName
    txtConfName.Enabled = chkCreateConf.Value
    
    Set btnApply = Me.Controls.Add("Forms.CommandButton.1")
    btnApply.Caption = "Apply"
    btnApply.Name = "btnApply"
    btnApply.Top = txtConfName.Top + txtConfName.height + MARGIN
    btnApply.Left = (maxWidth + MARGIN + TEXT_BOX_WIDTH) / 2 - btnApply.Width / 2 + MARGIN
    
    Dim height As Integer
    height = btnApply.Top + btnApply.height + MARGIN
    
    Me.StartUpPosition = 1 'center owner
    Me.ScrollBars = IIf(height > MAX_FORM_HEIGHT, fmScrollBarsVertical, fmScrollBarsNone)
    Me.ScrollHeight = height
    Me.Width = (maxWidth + MARGIN + TEXT_BOX_WIDTH) + MARGIN * 2 + 20
    Me.height = IIf(height > MAX_FORM_HEIGHT, MAX_FORM_HEIGHT + 25, height + 25) 'including header height
       
End Sub

Private Sub chkCreateConf_Change()
    txtConfName.Enabled = chkCreateConf.Value
End Sub

Private Sub btnApply_Click()
    
    Dim targConfName As String
    
    If chkCreateConf.Value Then
        targConfName = txtConfName.Text
    Else
        targConfName = FeatConfName
    End If
    
    Dim i As Integer
    
    For i = 0 To UBound(FeatDimsInfos)
        Dim dimValTxt As String
        dimValTxt = txtParamValue(i).Text
        If IsNumeric(dimValTxt) Then
            FeatDimsInfos(i).Value = CDbl(dimValTxt)
        Else
            Err.Raise vbError, "", "Specified value for " & FeatDimsInfos(i).title & " is not numeric"
        End If
    Next
    
    TrySetDimensions FeatDimsInfos, swActiveModel, targConfName, chkCreateConf.Value
    
End Sub