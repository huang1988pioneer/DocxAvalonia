using System;
using Avalonia;

namespace DocxAvalonia;

sealed class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        CrashLog.InstallGlobalHandlers();

        try
        {
            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            CrashLog.Write("Main fatal", ex);
            throw;
        }
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
