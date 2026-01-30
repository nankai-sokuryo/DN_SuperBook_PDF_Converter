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
            base.OnStartup(e);
            
            CoresLib.Init(new CoresLibOptions(CoresMode.Application, "SuperBookToolsGui", DebugMode.Debug, 
                defaultPrintStatToConsole: false, defaultRecordLeakFullStack: false));
        }

        protected override void OnExit(ExitEventArgs e)
        {
            CoresLib.Free();
            base.OnExit(e);
        }
    }
}
