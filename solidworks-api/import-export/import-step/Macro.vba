Option Explicit

'Overwrites if solidworks files already exist in case they have been processed before.
Const OVERWRITE As Boolean = False

'set the path you want to save to
Const DESTINATION_PATH As String = "C:\temp"


Sub main()

try_:

    'Uncomment the following line if you want to debug into this code during running Batch+
    'Debug.Assert False
    
    On Error GoTo catch_
    
    'test if DESTINATION PATH exists
    If FolderExists(DESTINATION_PATH) Then

        Dim swApp As SldWorks.SldWorks
        Set swApp = Application.SldWorks
        
        'You have to open a step file first, without saving it if you want to test without Batch+
        Dim swModel As SldWorks.ModelDoc2
        Set swModel = swApp.ActiveDoc
        
        If Not swModel Is Nothing Then
                    
             '--- Get file name without extension and path
             'only get the document name (which is displayed in the title bar of SolidWorks)
             Dim swxFilenaam As String
             swxFilenaam = swModel.GetTitle
             
             '--- Get file extension
             'Determine if the step file was an assembly or a part file to set the file extension correctly
             Dim Extension As String
             Select Case swModel.GetType
                
                Case swDocPART:
                    Extension = ".SLDPRT"
                
                Case swDocASSEMBLY:
                    Extension = ".SLDASM"
                    
             End Select
            
            '--- Get path
             Dim newPath As String
             newPath = DESTINATION_PATH
          
             
            'Add the name of the subfolder
             Dim subfoldername As String            
             subfoldername = "\" + swxFilenaam + "\"
             newPath = DESTINATION_PATH + subfoldername    
            
            '--- if folder doesn't exist already create it
             CreateFolderIfNotExisting (newPath)
            
            '--- Create the name of the file to save to
            swxFilenaam = newPath + swxFilenaam + Extension
            
            '--- if swxFilenaam exists already and OVERWRITE = False
            If FileExists(swxFilenaam) And OVERWRITE = False Then
                'do nothing
            Else
        
                ' make sure nothing is selected, otherwise only selected entities are saved
                swModel.ClearSelection2 False
        
        '--- save the step file
                Dim lErrors As Long
                Dim lWarnings As Long
                Dim boolstatus As Boolean
                boolstatus = swModel.Extension.SaveAs(swxFilenaam, 0, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, lErrors, lWarnings)
                Debug.Assert boolstatus
                                      
                'swApp.CloseDoc (swxFilenaam)'don't use it , let Batch+ handle it
             
             End If 'File exists already
             
        Else
            
            MsgBox "No document open"
            
        End If 'swModel Nothing
    
    Else
    
        MsgBox DESTINATION_PATH + "doesn't exist"
        
    End If 'DESTINATION_PATH exists
    
catch_:

    Debug.Print "Error: " & Err.Number & ":" & Err.source & ":" & Err.Description
    GoTo finally_
    
finally_:
    Debug.Print "FINISHED MACRO ImportStep"
    
End Sub

Function CreateFolderIfNotExisting(newPath As String)

    If FolderExists(newPath) Then
         'do nothing
    Else
        MkDir (newPath)
        Debug.Print "Path created : " + newPath
    End If

End Function

Function FolderExists(newPath As String) As Boolean

    If Dir(newPath, vbDirectory) = "" Then
        Debug.Print "Path doesn't exist : " + newPath
        FolderExists = False
    Else
        Debug.Print "Path exists : " + newPath
        FolderExists = True
    End If

End Function

Function FileExists(newPath As String) As Boolean

    If Dir(newPath) = "" Then
        Debug.Print "File doesn't exist : " + newPath
        FileExists = False
    Else
        Debug.Print "File exists : " + newPath
        FileExists = True
    End If

End Function
