Const DENOMINATOR As Integer = 16
Const ROUND_TO_NEAREST_FRACTION As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Debug.Print ConvertMetersToFractionInches(0.112713, DENOMINATOR, ROUND_TO_NEAREST_FRACTION)
    
End Sub

Function ConvertMetersToFractionInches(value As Double, denom As Integer, round As Boolean) As String
    
    Dim swUserUnits As SldWorks.UserUnit
    Set swUserUnits = swApp.GetUserUnit(swUserUnitsType_e.swLengthUnit)
    
    swUserUnits.FractionBase = swFractionDisplay_e.swFRACTION
    swUserUnits.SpecificUnitType = swLengthUnit_e.swINCHES
    
    swUserUnits.RoundToFraction = round
    swUserUnits.FractionValue = denom

    ConvertMetersToFractionInches = swUserUnits.ConvertToUserUnit(value, True, True)
    
End Function