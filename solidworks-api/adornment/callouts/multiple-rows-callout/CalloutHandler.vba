Implements swCalloutHandler

Private Function SwCalloutHandler_OnStringValueChanged(ByVal pManipulator As Object, ByVal RowID As Long, ByVal Text As String) As Boolean

        MsgBox "Text changed at row " & RowID & ": " & Text
                
        SwCalloutHandler_OnStringValueChanged = True

End Function