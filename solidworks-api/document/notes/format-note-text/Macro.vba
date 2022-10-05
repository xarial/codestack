Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = swModel.SelectionManager
    
    Dim swNote As SldWorks.Note
    
    Set swNote = swModel.InsertNote("<FONT effect=U>First Line Underline" & vbLf & "<FONT style=B effect=RU>Second Line Bold" & vbLf & "<FONT style=RB><FONT style=I>Third Line Italic")
        
    Debug.Print swNote.GetText()
    Debug.Print swNote.PropertyLinkedText
    
End Sub