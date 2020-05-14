Public WithEvents Part As SldWorks.PartDoc
Public FilePath As String

Private Function Part_FileSaveAsNotify2(ByVal fileName As String) As Long
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = Part
    
    swModel.SetSaveAsFileName FilePath
    Part_FileSaveAsNotify2 = 1
    
End Function