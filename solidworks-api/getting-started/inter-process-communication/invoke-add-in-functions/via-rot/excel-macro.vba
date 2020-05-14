Function GetFacesCount() As Integer
    
    Dim geomObjFactory As Object
    Set geomObjFactory = CreateObject("GeometryHelper.ApiObjectFactory")
    Dim geomHelper As Object
    
    Set geomHelper = geomObjFactory.GetInstance(13004)
    GetFacesCount = geomHelper.GetFacesCount(0)
    
End Function