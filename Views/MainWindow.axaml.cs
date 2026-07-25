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
    /// <summary>Cell under the last right-click, so menu acts on the correct table position.</summary>
    private Control? _tableContextSource;

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

    private void OnImagePointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (DataContext is not MainViewModel vm || sender is not Control control)
            return;

        // Select image block for resize commands
        for (var c = control as Control; c is not null; c = c.GetVisualParent() as Control)
        {
            if (c.DataContext is ImageBlock image)
            {
                vm.SelectedBlock = image;
                vm.SelectedImage = image;
                e.Handled = true;
                return;
            }
        }
    }

    private void OnTableCellGotFocus(object? sender, GotFocusEventArgs e)
    {
        SelectCellFromControl(sender as Control);
    }

    private void OnTableCellPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        SelectCellFromControl(sender as Control);
    }

    private void OnTableCellContextRequested(object? sender, ContextRequestedEventArgs e)
    {
        // Right-click: select that cell so expand commands target the correct row/column.
        _tableContextSource = sender as Control;
        SelectCellFromControl(_tableContextSource);
    }

    private void OnTableContextRequested(object? sender, ContextRequestedEventArgs e)
    {
        _tableContextSource = sender as Control;
        // If right-click on table chrome (not a cell), keep previous selection or last cell.
        if (DataContext is MainViewModel vm && sender is Control control)
        {
            if (FindBlockDataContext(control) is TableBlock table)
                vm.SelectedBlock = table;
        }
    }

    private void OnTableContextMenuOpening(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        // Ensure selection is updated before any menu command runs.
        if (_tableContextSource is not null)
            SelectCellFromControl(_tableContextSource);
        else if (sender is ContextMenu { PlacementTarget: Control target })
            SelectCellFromControl(target);
    }

    private void OnTableMenuAddRowBelow(object? sender, RoutedEventArgs e) =>
        RunTableCommand(vm => vm.TableAddRowBelowCommand.Execute(null));

    private void OnTableMenuAddColumnRight(object? sender, RoutedEventArgs e) =>
        RunTableCommand(vm => vm.TableAddColumnRightCommand.Execute(null));

    private void OnTableMenuInsertRowAbove(object? sender, RoutedEventArgs e) =>
        RunTableCommand(vm => vm.TableInsertRowAboveCommand.Execute(null));

    private void OnTableMenuInsertColumnLeft(object? sender, RoutedEventArgs e) =>
        RunTableCommand(vm => vm.TableInsertColumnLeftCommand.Execute(null));

    private void OnTableMenuDeleteRow(object? sender, RoutedEventArgs e) =>
        RunTableCommand(vm => vm.TableDeleteRowCommand.Execute(null));

    private void OnTableMenuDeleteColumn(object? sender, RoutedEventArgs e) =>
        RunTableCommand(vm => vm.TableDeleteColumnCommand.Execute(null));

    private void RunTableCommand(System.Action<MainViewModel> action)
    {
        if (_tableContextSource is not null)
            SelectCellFromControl(_tableContextSource);

        if (DataContext is MainViewModel vm)
            action(vm);
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
        else if (table is not null)
            vm.SelectedBlock = table;
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
