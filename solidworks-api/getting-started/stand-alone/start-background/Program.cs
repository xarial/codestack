private static ISldWorks StartSwAppBackground(string appPath, int timeoutSec = 20)
{
    var timeout = TimeSpan.FromSeconds(timeoutSec);

    var startTime = DateTime.Now;

    var prcInfo = new ProcessStartInfo()
    {
        FileName = appPath,
        Arguments = "/r", //no splash screen
        CreateNoWindow = true,
        WindowStyle = ProcessWindowStyle.Hidden
    };

    var prc = Process.Start(prcInfo);
    
    ISldWorks app = null;

    var isLoaded = false;

    var onIdleFunc = new DSldWorksEvents_OnIdleNotifyEventHandler(() =>
    {
        isLoaded = true;
        return 0;
    });

    try
    {

        while (!isLoaded)
        {
            if (DateTime.Now - startTime > timeout)
            {
                throw new TimeoutException();
            }

            if (app == null)
            {
                app = GetSwAppFromProcess(prc.Id);

                if (app != null)
                {
                    (app as SldWorks).OnIdleNotify += onIdleFunc;
                }
            }

            System.Threading.Thread.Sleep(100);
        }

        if (app != null)
        {
            const int HIDE = 0;
            ShowWindow(new IntPtr(app.IFrameObject().GetHWnd()), HIDE);
        }
    }
    catch
    {
        throw;
    }
    finally
    {
        if (app != null)
        {
            (app as SldWorks).OnIdleNotify -= onIdleFunc;
        }
    }

    return app;
}