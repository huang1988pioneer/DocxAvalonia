using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using DocxAvalonia.ViewModels;
using DocxAvalonia.Views;

namespace DocxAvalonia;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var args = desktop.Args ?? Array.Empty<string>();
            var initialPath = args.FirstOrDefault(a =>
                !string.IsNullOrWhiteSpace(a)
                && a.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));

            var vm = new MainViewModel();
            desktop.MainWindow = new MainWindow
            {
                DataContext = vm,
            };

            // 支援：WordViewHost.exe "C:\path\file.docx"
            if (initialPath is not null)
            {
                desktop.MainWindow.Opened += async (_, _) =>
                {
                    await vm.LoadDocumentAsync(initialPath);
                };
            }
        }

        base.OnFrameworkInitializationCompleted();
    }
}
