using System.Linq;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using DocxAvalonia.Models;
using DocxAvalonia.ViewModels;

namespace DocxAvalonia.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        AddHandler(DragDrop.DropEvent, OnDrop);
        AddHandler(DragDrop.DragOverEvent, OnDragOver);
    }

    private void OnDragOver(object? sender, DragEventArgs e)
    {
        e.DragEffects = HasDocxFile(e) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private async void OnDrop(object? sender, DragEventArgs e)
    {
        if (DataContext is not MainViewModel vm)
            return;

        var path = TryGetDocxPath(e);
        if (path is null)
            return;

        await vm.LoadDocumentAsync(path);
        e.Handled = true;
    }

    private static bool HasDocxFile(DragEventArgs e)
    {
        if (!e.Data.Contains(DataFormats.Files))
            return false;
        var files = e.Data.GetFiles();
        return files?.Any(f =>
            f.Name.EndsWith(".docx", System.StringComparison.OrdinalIgnoreCase)
            || (f.TryGetLocalPath()?.EndsWith(".docx", System.StringComparison.OrdinalIgnoreCase) ?? false)) == true;
    }

    private static string? TryGetDocxPath(DragEventArgs e)
    {
        var files = e.Data.GetFiles()?.ToList();
        if (files is null)
            return null;

        foreach (var item in files)
        {
            var local = item.TryGetLocalPath();
            if (!string.IsNullOrWhiteSpace(local)
                && local.EndsWith(".docx", System.StringComparison.OrdinalIgnoreCase))
                return local;
        }

        return null;
    }

    private void OnEditorGotFocus(object? sender, GotFocusEventArgs e)
    {
        if (DataContext is not MainViewModel vm || sender is not Control control)
            return;

        // Focusing a paragraph clears cell selection.
        vm.SelectedTableCell = null;
        var block = FindBlockDataContext(control);
        if (block is not null)
            vm.SelectedBlock = block;
    }

    private void OnBlockPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (DataContext is not MainViewModel vm || sender is not Control control)
            return;

        var block = FindBlockDataContext(control);
        if (block is not null)
            vm.SelectedBlock = block;
    }

    private void OnTableCellGotFocus(object? sender, GotFocusEventArgs e)
    {
        SelectCellFromControl(sender as Control);
    }

    private void OnTableCellPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        SelectCellFromControl(sender as Control);
    }

    private void SelectCellFromControl(Control? control)
    {
        if (DataContext is not MainViewModel vm || control is null)
            return;

        TableCellBlock? cell = null;
        TableBlock? table = null;
        for (var c = control as Control; c is not null; c = c.GetVisualParent() as Control)
        {
            if (cell is null && c.DataContext is TableCellBlock cellBlock)
                cell = cellBlock;
            if (c.DataContext is TableBlock tableBlock)
            {
                table = tableBlock;
                break;
            }
        }

        if (cell is not null)
            vm.SelectTableCell(cell, table);
    }

    private static DocumentBlock? FindBlockDataContext(Control control)
    {
        for (var c = control as Control; c is not null; c = c.GetVisualParent() as Control)
        {
            if (c.DataContext is DocumentBlock block)
                return block;
        }

        return control.DataContext as DocumentBlock;
    }
}
