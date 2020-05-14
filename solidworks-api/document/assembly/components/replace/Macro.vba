Const REPLACEMENT_DIR As String = "D:\Assembly\Replacement"
Const SUFFIX As String = "_new"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim swAssy As SldWorks.AssemblyDoc
        Set swAssy = swModel
        
        Dim swSelMgr As SldWorks.SelectionMgr
        Set swSelMgr = swModel.SelectionManager
        
        Dim i As Integer
        
        For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
            
            If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelCOMPONENTS Then
                
                Dim swComp As SldWorks.Component2
                Set swComp = swSelMgr.GetSelectedObject6(i, -1)
                
                Debug.Print swSelMgr.SuspendSelectionList
                
                swSelMgr.AddSelectionListObject swComp, Nothing
                
                swAssy.ReplaceComponents2 GetReplacementPath(swComp), swComp.ReferencedConfiguration, False, swReplaceComponentsConfiguration_e.swReplaceComponentsConfiguration_MatchName, True
                    
                swSelMgr.ResumeSelectionList
                
            End If
        Next
        
    Else
        MsgBox ("Please open assembly document")
    End If
    
End Sub

Function GetReplacementPath(comp As SldWorks.Component2)
    
    Dim replFilePath As String
    
    Dim compPath As String
    compPath = comp.GetPathName()
                
    Dim dir As String
    dir = REPLACEMENT_DIR
    
    If Right(dir, 1) <> "\" Then
        dir = dir & "\"
    End If
    
    Dim fileName As String
    fileName = Right(compPath, Len(compPath) - InStrRev(compPath, "\"))
    
    If SUFFIX <> "" Then
        
        Dim ext As String
        
        ext = Right(fileName, Len(".SLDXXX"))
        
        fileName = Left(fileName, Len(fileName) - Len(ext)) & SUFFIX & ext
        
    End If
    
    replFilePath = dir & fileName
                
    GetReplacementPath = replFilePath
    
End Function