Sub main()

    Dim userType As MyUserType
    Set userType = New MyUserType
    Dim obj As Object
    
    'obj = userType 'Object variable or with block variable not set when Set keyword is not used
    Set obj = userType 'assigning the pointer to the MyUserType object to obj variable

End Sub