Private Async Function StartSwAppAsync(ByVal appPath As String, _
    ByVal Optional timeoutSec As Integer = 10) _
        As System.Threading.Tasks.Task(Of SolidWorks.Interop.sldworks.ISldWorks)
    Return Await System.Threading.Tasks.Task.Run(Function() StartSwApp(appPath, timeoutSec))
End Function
