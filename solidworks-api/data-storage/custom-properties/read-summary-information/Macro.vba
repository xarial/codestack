Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        Debug.Print "Author: " & swModel.SummaryInfo(swSummInfoField_e.swSumInfoAuthor)
        Debug.Print "Keywords: " & swModel.SummaryInfo(swSummInfoField_e.swSumInfoKeywords)
        Debug.Print "Comments: " & swModel.SummaryInfo(swSummInfoField_e.swSumInfoComment)
        Debug.Print "Title: " & swModel.SummaryInfo(swSummInfoField_e.swSumInfoTitle)
        Debug.Print "Subject: " & swModel.SummaryInfo(swSummInfoField_e.swSumInfoSubject)
        
        Debug.Print "Created: " & swModel.SummaryInfo(swSummInfoField_e.swSumInfoCreateDate2)
        Debug.Print "Last Saved: " & swModel.SummaryInfo(swSummInfoField_e.swSumInfoSaveDate2)
        Debug.Print "Last Saved By: " & swModel.SummaryInfo(swSummInfoField_e.swSumInfoSavedBy)
        Debug.Print "Last Saved With: " & GetLastSavedWith(swModel)
    Else
        MsgBox "Please open model"
    End If
    
End Sub

Function GetLastSavedWith(model As SldWorks.ModelDoc2) As String
    
    Dim vHistory As Variant
    vHistory = model.VersionHistory()
    
    Dim vers As String
    vers = vHistory(UBound(vHistory))
    
    vers = Mid(vers, InStr(vers, "[") + 1, 4)
    
    GetLastSavedWith = "SOLIDWORKS " & vers
    
End Function