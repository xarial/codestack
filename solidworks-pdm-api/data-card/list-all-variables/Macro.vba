Sub main()

    Dim swPdmVault As New EdmVault5
    swPdmVault.LoginAuto "TestVault", 0
    
    If swPdmVault.IsLoggedIn Then
        
        Dim swPdmVarsMgr As IEdmVariableMgr7
        Set swPdmVarsMgr = swPdmVault
        
        Dim swVarPost As IEdmPos5
        Set swVarPost = swPdmVarsMgr.GetFirstVariablePosition()
        
        While Not swVarPost.IsNull
            Dim swPdmVar As IEdmVariable5
            Set swPdmVar = swPdmVarsMgr.GetNextVariable(swVarPost)
            Debug.Print swPdmVar.Name & "(" & swPdmVar.ID & ")"
        Wend
    Else
        Err.Raise vberr, "", "Not logged in"
    End If

End Sub