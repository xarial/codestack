Public Sub GetTolerance(ByVal [dim] As IDimension)
    Dim dimTol = [dim].Tolerance
    Dim maxTol As Double
    Dim minTol As Double

    If App.IsVersionNewerOrEqual(SwVersion_e.Sw2015, 3) Then
        dimTol.GetMinValue2(minTol)
        dimTol.GetMaxValue2(maxTol)
    Else
        minTol = dimTol.GetMinValue
        maxTol = dimTol.GetMaxValue
    End If
End Sub