#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as Long) 'For 32 Bit Systems
#End If

Sub main()

    Debug.Print Now
    Wait 20000, False
    Debug.Print Now

End Sub

Sub Wait(period As Long, blocked As Boolean)

    If blocked Then
        Sleep period
    Else
        Const STEP As Long = 100
        
        If period > STEP Then
            
            Dim i As Long
            
            For i = 0 To period Step STEP
                Sleep STEP
                DoEvents
            Next
            
        Else
            Sleep period
        End If
        
    End If

End Sub