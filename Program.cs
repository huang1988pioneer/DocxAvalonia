using System;
using System.Diagnostics;
using Avalonia;

namespace DocxAvalonia;

sealed class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        // Prevent one-off image/decode faults from taking down the whole process when possible.
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            Debug.WriteLine(e.ExceptionObject);
        };
        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            Debug.WriteLine(e.Exception);
            e.SetObserved();
        };

        try
        {
            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex);
            throw;
        }
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
