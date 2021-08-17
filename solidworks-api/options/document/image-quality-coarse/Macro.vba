Option Explicit

' An assembly or a part file must be the active document.

' the document options are set to use coarse quality
' and the checkmark "Apply to all referenced part documents" is set to ON if the active document is an assembly

Dim swxApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2


Sub main()

try_:

    On Error GoTo catch_

    Set swxApp = Application.SldWorks
    
    Set swModel = swxApp.ActiveDoc

    'Check if active document is a Part or an Assembly file
    Select Case True
    
           Case swModel Is Nothing, (swModel.GetType <> swDocASSEMBLY And swModel.GetType <> swDocPART)
              Call swxApp.SendMsgToUser2("Please open an assembly or part file", swMbInformation, swMbOk)
                           
           Case Else
               Call SetCoarseQuality
               
    End Select

    GoTo finally_:
    
catch_:

        Debug.Print "Error: " & Err.Number & ":" & Err.Source & ":" & Err.Description
    
finally_:
    
End Sub

Private Function SetCoarseQuality() As Boolean
                  
      'set to use coarse quality
      Dim boolstatus As Boolean
      boolstatus = swModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swImageQualityShaded, _
                                                              swUserPreferenceOption_e.swDetailingNoOptionSpecified, _
                                                              swImageQualityShaded_e.swShadedImageQualityCoarse)
        
      'option "Apply to all referenced part documents" is set to ON
      If swModel.GetType = swDocASSEMBLY Then
      
         Dim res As Boolean
         res = swModel.Extension.SetUserPreferenceToggle(swImageQualityApplyToAllReferencedPartDoc, _
                                                         swDetailingNoOptionSpecified, True)
        
      End If
           
End Function








