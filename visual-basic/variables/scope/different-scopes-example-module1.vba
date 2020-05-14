Public Module1PublicText As String 'public variable is visible outside of the module
Dim Module1PrivateText As String 'only visible by functions and properties of this module

Sub Init()
    
    Module1PublicText = "Module1 Public Text"
    Module1PrivateText = "Module1 Private Text"
    
End Sub