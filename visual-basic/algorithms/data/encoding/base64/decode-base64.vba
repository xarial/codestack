Private Function Base64ToArray(base64 As String) As Variant
    
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.Text = base64
    
    Base64ToArray = xmlNode.nodeTypedValue
  
End Function