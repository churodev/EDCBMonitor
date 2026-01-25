using System.Text;
using System.Threading;
using System.Windows;

namespace EDCBMonitor;

// System.Windows.Application を明示的に継承
public partial class App : System.Windows.Application
{
    private Mutex? _mutex;

    public App()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        _mutex = new Mutex(true, "Global\\EDCBMonitor_Mutex_Net8", out bool createdNew);

        if (!createdNew)
        {
            _mutex = null;
            Shutdown();
            return;
        }

        Config.Load();
        base.OnStartup(e);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        if (_mutex != null)
        {
            _mutex.ReleaseMutex();
            _mutex.Dispose();
        }
        Config.Save();
        base.OnExit(e);
    }
}