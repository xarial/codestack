<Title(GetType(Resources), NameOf(Resources.MacroFeatureBaseName))>
<ComVisible(True)>
Public Class LocalizedMacroFeature
    Inherits MacroFeatureEx

    Protected Overrides Function OnRebuild(ByVal app As ISldWorks, ByVal model As IModelDoc2, ByVal feature As IFeature) As MacroFeatureRebuildResult
        If Not String.IsNullOrEmpty(model.GetPathName) Then
            Return MacroFeatureRebuildResult.FromStatus(True)
        Else
            Return MacroFeatureRebuildResult.FromStatus(False, Resources.MacroFeatureError)
        End If
    End Function
End Class