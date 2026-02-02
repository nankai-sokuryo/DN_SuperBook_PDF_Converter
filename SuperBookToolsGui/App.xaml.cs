using System;
using System.Windows;

using IPA.Cores.Basic;
using IPA.Cores.Helper.Basic;
using static IPA.Cores.Globals.Basic;

namespace SuperBookToolsGui
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Ensure application exits when main window closes
            this.ShutdownMode = ShutdownMode.OnMainWindowClose;
            
            base.OnStartup(e);
            
            CoresLib.Init(new CoresLibOptions(CoresMode.Application, "SuperBookToolsGui", DebugMode.Debug, 
                defaultPrintStatToConsole: false, defaultRecordLeakFullStack: false));
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // CoresLib.Free() is already called in MainWindow.StartShutdownAsync()
            // No need to call it again here
            base.OnExit(e);
        }
    }
}
