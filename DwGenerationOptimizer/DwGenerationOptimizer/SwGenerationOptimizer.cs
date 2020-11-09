using System;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using SolidWorks.Interop.sldworks;

namespace BlueGiant
{    
    public class SwGenerationOptimizer
    {
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("ole32.dll")]
        private static extern void CreateBindCtx(int reserved, out IBindCtx ppbc);

        static void Main(string[] args)
        {
            const string SW_PATH = @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe";

            try
            {
                var app = StartSwAppGenerationOptimized(SW_PATH);
                Console.WriteLine("BlueGiant.SwGenerationOptimizer - SolidWorks running in backgroud and ready for generation");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to connect to SolidWorks instance: " + ex.Message);
            }

            Console.ReadLine();
        }

        private static ISldWorks StartSwAppGenerationOptimized(string appPath, int timeoutSec = 20)
        {
            var timeout = TimeSpan.FromSeconds(timeoutSec);
            var startTime = DateTime.Now;

            var prcInfo = new ProcessStartInfo()
            {
                FileName = appPath,
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

        private static ISldWorks GetSwAppFromProcess(int processId)
        {
            var monikerName = "SolidWorks_PID_" + processId.ToString();

            IBindCtx context = null;
            IRunningObjectTable rot = null;
            IEnumMoniker monikers = null;

            try
            {
                CreateBindCtx(0, out context);

                context.GetRunningObjectTable(out rot);
                rot.EnumRunning(out monikers);

                var moniker = new IMoniker[1];

                while (monikers.Next(1, moniker, IntPtr.Zero) == 0)
                {
                    var curMoniker = moniker.First();

                    string name = null;

                    if (curMoniker != null)
                    {
                        try
                        {
                            curMoniker.GetDisplayName(context, null, out name);
                        }
                        catch (UnauthorizedAccessException)
                        {
                        }
                    }

                    if (string.Equals(monikerName, name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        object app;
                        rot.GetObject(curMoniker, out app);
                        return app as ISldWorks;
                    }
                }
            }
            finally
            {
                if (monikers != null)
                {
                    Marshal.ReleaseComObject(monikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (context != null)
                {
                    Marshal.ReleaseComObject(context);
                }
            }

            return null;
        }
    }
}