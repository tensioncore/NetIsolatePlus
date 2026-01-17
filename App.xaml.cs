using System;
using System.Windows;
using System.Windows.Threading;

namespace NetIsolatePlus;

public partial class App : Application
{
    public App()
    {
        this.DispatcherUnhandledException += (s, e) =>
        {
            try
            {
                if (Current?.Dispatcher?.HasShutdownStarted == true) return;

                MessageBox.Show("Unhandled error:\n\n" + e.Exception, "NetIsolate+",
                    MessageBoxButton.OK, MessageBoxImage.Error);

                e.Handled = true; // prevent silent exit
            }
            catch
            {
                // last resort: let it crash if we can't show UI
            }
        };

        AppDomain.CurrentDomain.UnhandledException += (s, e) =>
        {
            try
            {
                var ex = e.ExceptionObject as Exception;
                var msg = "Fatal error:\n\n" + ex;

                // This can fire on a non-UI thread; only attempt UI marshal (avoid blocking/hanging here).
                if (Current?.Dispatcher != null && !Current.Dispatcher.HasShutdownStarted)
                {
                    Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            if (Current?.Dispatcher?.HasShutdownStarted == true) return;

                            MessageBox.Show(msg, "NetIsolate+",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        catch { }
                    }), DispatcherPriority.Send);
                }
            }
            catch
            {
                // swallow; nothing reliable to do here
            }
        };
    }
}
