using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input.Platform;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocxAvalonia.Models;
using DocxAvalonia.Services;

namespace DocxAvalonia.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    public const int AutoSaveIntervalSeconds = 60;

    private readonly DocxDocumentService _documentService = new();
    private readonly Stack<string> _undoStack = new();
    private readonly Stack<string> _redoStack = new();
    private readonly DispatcherTimer _autoSaveTimer;
    private bool _suspendDirty;

    [ObservableProperty]
    private string _windowTitle = "DocxAvalonia";

    [ObservableProperty]
    private string _statusText = "就緒 — 新增或開啟文件開始編輯";

    [ObservableProperty]
    private string? _filePath;

    [ObservableProperty]
    private string? _fileName;

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private bool _hasDocument;

    [ObservableProperty]
    private bool _isDirty;

    /// <summary>When true, dirty documents with a known path are saved every 60 seconds.</summary>
    [ObservableProperty]
    private bool _autoSaveEnabled;

    [ObservableProperty]
    private string _autoSaveStatusText = "自動儲存：關";

    [ObservableProperty]
    private WordDocument? _document;

    [ObservableProperty]
    private DocumentBlock? _selectedBlock;

    [ObservableProperty]
    private ParagraphBlock? _selectedParagraph;

    [ObservableProperty]
    private double _zoom = 1.0;

    [ObservableProperty]
    private string _zoomText = "100%";

    [ObservableProperty]
    private string _statisticsText = "字數: 0";

    [ObservableProperty]
    private string _findText = string.Empty;

    [ObservableProperty]
    private string _replaceText = string.Empty;

    [ObservableProperty]
    private bool _showFindPanel;

    [ObservableProperty]
    private int _findMatchCount;

    public ObservableCollection<double> FontSizeOptions { get; } =
        [10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48];

    public MainViewModel()
    {
        _autoSaveTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(AutoSaveIntervalSeconds),
        };
        _autoSaveTimer.Tick += OnAutoSaveTick;

        NewDocument();
    }

    partial void OnZoomChanged(double value)
    {
        ZoomText = $"{(int)Math.Round(value * 100)}%";
    }

    partial void OnAutoSaveEnabledChanged(bool value)
    {
        if (value)
        {
            _autoSaveTimer.Start();
            AutoSaveStatusText = $"自動儲存：每 {AutoSaveIntervalSeconds} 秒";
            StatusText = $"已開啟自動儲存（每 {AutoSaveIntervalSeconds} 秒；需已有儲存路徑）";
        }
        else
        {
            _autoSaveTimer.Stop();
            AutoSaveStatusText = "自動儲存：關";
            StatusText = "已關閉自動儲存";
        }
    }

    private async void OnAutoSaveTick(object? sender, EventArgs e)
    {
        if (!AutoSaveEnabled || Document is null || IsBusy || !IsDirty)
            return;

        if (string.IsNullOrWhiteSpace(FilePath))
        {
            StatusText = "自動儲存略過：請先「另存新檔」指定路徑";
            return;
        }

        await SaveToPathAsync(FilePath, isAutoSave: true);
    }

    [RelayCommand]
    private void ToggleAutoSave()
    {
        AutoSaveEnabled = !AutoSaveEnabled;
    }

    partial void OnSelectedBlockChanged(DocumentBlock? value)
    {
        SelectedParagraph = value as ParagraphBlock;
    }

    partial void OnDocumentChanged(WordDocument? value)
    {
        HasDocument = value is not null;
        RefreshStatistics();
        UpdateTitle();
    }

    partial void OnIsDirtyChanged(bool value) => UpdateTitle();

    partial void OnFileNameChanged(string? value) => UpdateTitle();

    private void UpdateTitle()
    {
        var name = string.IsNullOrWhiteSpace(FileName) ? "未命名文件" : FileName;
        var dirty = IsDirty ? " *" : string.Empty;
        WindowTitle = $"{name}{dirty} — DocxAvalonia";
    }

    private void RefreshStatistics()
    {
        if (Document is null)
        {
            StatisticsText = "字數: 0";
            return;
        }

        var (words, chars, charsNoSpaces, paragraphs) = Document.GetStatistics();
        StatisticsText = $"字數: {words}  |  字元: {chars}（不含空白 {charsNoSpaces}）  |  段落: {paragraphs}";
    }

    private void MarkDirty()
    {
        if (_suspendDirty)
            return;
        IsDirty = true;
        RefreshStatistics();
    }

    private void PushUndo()
    {
        if (Document is null || _suspendDirty)
            return;
        try
        {
            _undoStack.Push(SerializeDocument(Document));
            if (_undoStack.Count > 40)
            {
                var arr = _undoStack.Reverse().Take(40).Reverse().ToArray();
                _undoStack.Clear();
                foreach (var s in arr)
                    _undoStack.Push(s);
            }

            _redoStack.Clear();
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex);
        }
    }

    private static string SerializeDocument(WordDocument doc)
    {
        var dto = new DocDto
        {
            Title = doc.Title,
            Blocks = doc.Blocks.Select(b => b switch
            {
                ParagraphBlock p => new BlockDto
                {
                    Type = "p",
                    Text = p.Text,
                    Style = p.Style.ToString(),
                    Align = p.Alignment.ToString(),
                    Bold = p.IsBold,
                    Italic = p.IsItalic,
                    Underline = p.IsUnderline,
                    FontSize = p.FontSize,
                    List = p.ListKind.ToString(),
                },
                TableBlock t => new BlockDto
                {
                    Type = "t",
                    Rows = t.Rows.Select(r => r.Cells.Select(c => c.Text).ToList()).ToList(),
                },
                ImageBlock img => new BlockDto
                {
                    Type = "i",
                    Text = img.FileName,
                    ContentType = img.ContentType,
                    ImageBase64 = Convert.ToBase64String(img.ImageBytes),
                    DisplayWidth = img.DisplayWidth,
                    DisplayHeight = img.DisplayHeight,
                },
                _ => new BlockDto { Type = "p" },
            }).ToList(),
        };
        return JsonSerializer.Serialize(dto);
    }

    private static WordDocument DeserializeDocument(string json)
    {
        var dto = JsonSerializer.Deserialize<DocDto>(json)
            ?? throw new InvalidOperationException("無法還原文件。");
        var doc = new WordDocument { Title = dto.Title ?? "文件" };
        foreach (var b in dto.Blocks ?? [])
        {
            if (b.Type == "t" && b.Rows is not null)
            {
                var table = new TableBlock();
                foreach (var row in b.Rows)
                {
                    var tr = new TableRowBlock();
                    foreach (var cell in row)
                        tr.Cells.Add(new TableCellBlock { Text = cell });
                    table.Rows.Add(tr);
                }

                doc.Blocks.Add(table);
            }
            else if (b.Type == "i" && !string.IsNullOrEmpty(b.ImageBase64))
            {
                doc.Blocks.Add(new ImageBlock
                {
                    ImageBytes = Convert.FromBase64String(b.ImageBase64),
                    ContentType = b.ContentType ?? "image/png",
                    FileName = b.Text,
                    DisplayWidth = b.DisplayWidth > 0 ? b.DisplayWidth : 400,
                    DisplayHeight = b.DisplayHeight,
                });
            }
            else
            {
                doc.Blocks.Add(new ParagraphBlock
                {
                    Text = b.Text ?? string.Empty,
                    Style = Enum.TryParse<ParagraphStyleKind>(b.Style, out var s) ? s : ParagraphStyleKind.Normal,
                    Alignment = Enum.TryParse<ParagraphAlignmentKind>(b.Align, out var a) ? a : ParagraphAlignmentKind.Left,
                    IsBold = b.Bold,
                    IsItalic = b.Italic,
                    IsUnderline = b.Underline,
                    FontSize = b.FontSize <= 0 ? 14 : b.FontSize,
                    ListKind = Enum.TryParse<ListKind>(b.List, out var l) ? l : ListKind.None,
                });
            }
        }

        if (doc.Blocks.Count == 0)
            doc.Blocks.Add(new ParagraphBlock());
        return doc;
    }

    private void AttachDocument(WordDocument doc, string? path, bool dirty)
    {
        DetachDocumentHandlers();
        Document = doc;
        FilePath = path;
        FileName = path is null ? "未命名文件.docx" : Path.GetFileName(path);
        IsDirty = dirty;
        SelectedBlock = doc.Blocks.FirstOrDefault();
        _undoStack.Clear();
        _redoStack.Clear();
        AttachDocumentHandlers();
        RefreshStatistics();
        StatusText = path is null ? "已建立新文件" : $"已開啟：{FileName}";
    }

    private void AttachDocumentHandlers()
    {
        if (Document is null)
            return;
        Document.Blocks.CollectionChanged += OnBlocksChanged;
        foreach (var block in Document.Blocks)
            HookBlock(block);
    }

    private void DetachDocumentHandlers()
    {
        if (Document is null)
            return;
        Document.Blocks.CollectionChanged -= OnBlocksChanged;
        foreach (var block in Document.Blocks)
            UnhookBlock(block);
    }

    private void OnBlocksChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems)
            {
                if (item is DocumentBlock b)
                    HookBlock(b);
            }
        }

        MarkDirty();
        RefreshStatistics();
    }

    private void HookBlock(DocumentBlock block)
    {
        block.PropertyChanged += OnBlockPropertyChanged;
        if (block is TableBlock table)
        {
            foreach (var row in table.Rows)
            foreach (var cell in row.Cells)
                cell.PropertyChanged += OnBlockPropertyChanged;
        }
    }

    private void UnhookBlock(DocumentBlock block)
    {
        block.PropertyChanged -= OnBlockPropertyChanged;
        if (block is TableBlock table)
        {
            foreach (var row in table.Rows)
            foreach (var cell in row.Cells)
                cell.PropertyChanged -= OnBlockPropertyChanged;
        }
    }

    private void OnBlockPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        MarkDirty();
        if (e.PropertyName is nameof(ParagraphBlock.Text) or nameof(TableCellBlock.Text))
            RefreshStatistics();
    }

    // ─── File ───────────────────────────────────────────────

    [RelayCommand]
    private void NewDocument()
    {
        if (!ConfirmDiscardChanges())
            return;
        AttachDocument(WordDocument.CreateBlank(), null, dirty: false);
        StatusText = "已建立新文件";
    }

    [RelayCommand]
    private async Task OpenDocumentAsync()
    {
        if (!ConfirmDiscardChanges())
            return;

        var window = GetMainWindow();
        if (window is null)
            return;

        var files = await window.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "開啟 Word 文件",
            AllowMultiple = false,
            FileTypeFilter = [DocxFileType(), AllFileType()],
        });
        if (files.Count == 0)
            return;

        var path = files[0].TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path))
        {
            StatusText = "無法讀取檔案路徑。";
            return;
        }

        await LoadDocumentAsync(path);
    }

    public async Task LoadDocumentAsync(string path)
    {
        if (IsBusy)
            return;

        IsBusy = true;
        StatusText = "正在開啟…";
        try
        {
            // Parse OpenXML off UI thread (bytes only — no Avalonia Bitmap yet).
            var doc = await Task.Run(() => _documentService.Load(path));

            // Decode previews on UI thread to avoid Skia/threading crashes.
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                foreach (var image in doc.Blocks.OfType<ImageBlock>())
                {
                    try
                    {
                        image.EnsurePreview();
                    }
                    catch
                    {
                        // Keep document open even if one image fails.
                    }
                }
            });

            AttachDocument(doc, path, dirty: false);

            var imageCount = doc.Blocks.OfType<ImageBlock>().Count();
            var shown = doc.Blocks.OfType<ImageBlock>().Count(i => i.HasPreview);
            if (imageCount > 0)
                StatusText = $"已開啟：{FileName}（圖片 {shown}/{imageCount}）";
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex);
            StatusText = $"開啟失敗：{ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task SaveDocumentAsync()
    {
        if (Document is null)
            return;

        if (string.IsNullOrWhiteSpace(FilePath))
        {
            await SaveDocumentAsAsync();
            return;
        }

        await SaveToPathAsync(FilePath, isAutoSave: false);
    }

    [RelayCommand]
    private async Task SaveDocumentAsAsync()
    {
        if (Document is null)
            return;

        var window = GetMainWindow();
        if (window is null)
            return;

        var file = await window.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "另存新檔",
            SuggestedFileName = FileName ?? "未命名文件.docx",
            DefaultExtension = "docx",
            FileTypeChoices = [DocxFileType()],
        });
        if (file is null)
            return;

        var path = file.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path))
        {
            StatusText = "無法取得儲存路徑。";
            return;
        }

        if (!path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            path += ".docx";

        await SaveToPathAsync(path, isAutoSave: false);
    }

    private async Task SaveToPathAsync(string path, bool isAutoSave)
    {
        if (Document is null)
            return;

        IsBusy = true;
        if (!isAutoSave)
            StatusText = "正在儲存…";
        try
        {
            var doc = Document;
            await Task.Run(() => _documentService.Save(doc, path));
            FilePath = path;
            FileName = Path.GetFileName(path);
            IsDirty = false;
            var time = DateTime.Now.ToString("HH:mm:ss");
            StatusText = isAutoSave
                ? $"自動儲存完成 {time}：{FileName}"
                : $"已儲存：{FileName}";
        }
        catch (Exception ex)
        {
            StatusText = isAutoSave
                ? $"自動儲存失敗：{ex.Message}"
                : $"儲存失敗：{ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void CloseDocument()
    {
        if (!ConfirmDiscardChanges())
            return;
        NewDocument();
        StatusText = "已關閉文件";
    }

    private bool ConfirmDiscardChanges()
    {
        // No modal dialog service; for now auto-allow but warn in status if dirty.
        // A real confirm would use a dialog — keep flow non-blocking.
        if (IsDirty)
            StatusText = "注意：先前文件有未儲存變更，已建立/開啟新內容。";
        return true;
    }

    // ─── Edit ───────────────────────────────────────────────

    [RelayCommand]
    private void Undo()
    {
        if (Document is null || _undoStack.Count == 0)
            return;
        _redoStack.Push(SerializeDocument(Document));
        var json = _undoStack.Pop();
        _suspendDirty = true;
        try
        {
            var path = FilePath;
            var name = FileName;
            var restored = DeserializeDocument(json);
            DetachDocumentHandlers();
            Document = restored;
            FilePath = path;
            FileName = name;
            AttachDocumentHandlers();
            SelectedBlock = Document.Blocks.FirstOrDefault();
            IsDirty = true;
            RefreshStatistics();
            StatusText = "已復原";
        }
        finally
        {
            _suspendDirty = false;
        }
    }

    [RelayCommand]
    private void Redo()
    {
        if (Document is null || _redoStack.Count == 0)
            return;
        _undoStack.Push(SerializeDocument(Document));
        var json = _redoStack.Pop();
        _suspendDirty = true;
        try
        {
            var path = FilePath;
            var name = FileName;
            var restored = DeserializeDocument(json);
            DetachDocumentHandlers();
            Document = restored;
            FilePath = path;
            FileName = name;
            AttachDocumentHandlers();
            SelectedBlock = Document.Blocks.FirstOrDefault();
            IsDirty = true;
            RefreshStatistics();
            StatusText = "已取消復原";
        }
        finally
        {
            _suspendDirty = false;
        }
    }

    [RelayCommand]
    private async Task CutAsync()
    {
        await CopyAsync();
        if (SelectedParagraph is not null)
        {
            PushUndo();
            SelectedParagraph.Text = string.Empty;
            MarkDirty();
        }
    }

    [RelayCommand]
    private async Task CopyAsync()
    {
        var text = SelectedParagraph?.Text ?? Document?.GetPlainText() ?? string.Empty;
        var clipboard = GetClipboard();
        if (clipboard is null)
            return;
        await clipboard.SetTextAsync(text);
        StatusText = "已複製到剪貼簿";
    }

    [RelayCommand]
    private async Task PasteAsync()
    {
        var clipboard = GetClipboard();
        if (clipboard is null || SelectedParagraph is null)
            return;
        var text = await clipboard.GetTextAsync();
        if (string.IsNullOrEmpty(text))
            return;
        PushUndo();
        SelectedParagraph.Text = (SelectedParagraph.Text ?? string.Empty) + text;
        MarkDirty();
        StatusText = "已貼上";
    }

    [RelayCommand]
    private void SelectAllParagraphs()
    {
        if (Document?.Blocks.OfType<ParagraphBlock>().FirstOrDefault() is { } first)
            SelectedBlock = first;
        StatusText = "已選取文件內容（格式套用至目前段落）";
    }

    // ─── Format ─────────────────────────────────────────────

    [RelayCommand]
    private void ToggleBold()
    {
        if (SelectedParagraph is null)
            return;
        PushUndo();
        SelectedParagraph.IsBold = !SelectedParagraph.IsBold;
        MarkDirty();
    }

    [RelayCommand]
    private void ToggleItalic()
    {
        if (SelectedParagraph is null)
            return;
        PushUndo();
        SelectedParagraph.IsItalic = !SelectedParagraph.IsItalic;
        MarkDirty();
    }

    [RelayCommand]
    private void ToggleUnderline()
    {
        if (SelectedParagraph is null)
            return;
        PushUndo();
        SelectedParagraph.IsUnderline = !SelectedParagraph.IsUnderline;
        MarkDirty();
    }

    [RelayCommand]
    private void IncreaseFontSize()
    {
        if (SelectedParagraph is null)
            return;
        PushUndo();
        SelectedParagraph.FontSize = Math.Min(72, SelectedParagraph.FontSize + 2);
        MarkDirty();
    }

    [RelayCommand]
    private void DecreaseFontSize()
    {
        if (SelectedParagraph is null)
            return;
        PushUndo();
        SelectedParagraph.FontSize = Math.Max(8, SelectedParagraph.FontSize - 2);
        MarkDirty();
    }

    [RelayCommand]
    private void SetFontSize(double size)
    {
        if (SelectedParagraph is null)
            return;
        PushUndo();
        SelectedParagraph.FontSize = size;
        MarkDirty();
    }

    [RelayCommand]
    private void SetAlignment(string align)
    {
        if (SelectedParagraph is null)
            return;
        if (!Enum.TryParse<ParagraphAlignmentKind>(align, true, out var kind))
            return;
        PushUndo();
        SelectedParagraph.Alignment = kind;
        MarkDirty();
    }

    [RelayCommand]
    private void SetStyle(string style)
    {
        if (SelectedParagraph is null)
            return;
        if (!Enum.TryParse<ParagraphStyleKind>(style, true, out var kind))
            return;
        PushUndo();
        SelectedParagraph.Style = kind;
        MarkDirty();
    }

    [RelayCommand]
    private void ToggleBulletList()
    {
        if (SelectedParagraph is null)
            return;
        PushUndo();
        SelectedParagraph.ListKind = SelectedParagraph.ListKind == ListKind.Bullet
            ? ListKind.None
            : ListKind.Bullet;
        MarkDirty();
    }

    [RelayCommand]
    private void ToggleNumberedList()
    {
        if (SelectedParagraph is null)
            return;
        PushUndo();
        SelectedParagraph.ListKind = SelectedParagraph.ListKind == ListKind.Numbered
            ? ListKind.None
            : ListKind.Numbered;
        MarkDirty();
    }

    // ─── Insert ─────────────────────────────────────────────

    [RelayCommand]
    private void InsertParagraph()
    {
        if (Document is null)
            return;
        PushUndo();
        var p = new ParagraphBlock();
        var index = SelectedBlock is null ? Document.Blocks.Count : Document.Blocks.IndexOf(SelectedBlock) + 1;
        if (index < 0 || index > Document.Blocks.Count)
            index = Document.Blocks.Count;
        Document.Blocks.Insert(index, p);
        SelectedBlock = p;
        MarkDirty();
        StatusText = "已插入段落";
    }

    [RelayCommand]
    private void InsertTable()
    {
        if (Document is null)
            return;
        PushUndo();
        var table = TableBlock.Create(3, 3);
        var index = SelectedBlock is null ? Document.Blocks.Count : Document.Blocks.IndexOf(SelectedBlock) + 1;
        if (index < 0 || index > Document.Blocks.Count)
            index = Document.Blocks.Count;
        Document.Blocks.Insert(index, table);
        SelectedBlock = table;
        MarkDirty();
        StatusText = "已插入 3×3 表格";
    }

    [RelayCommand]
    private async Task InsertImageAsync()
    {
        if (Document is null)
            return;

        var window = GetMainWindow();
        if (window is null)
            return;

        var files = await window.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "插入圖片",
            AllowMultiple = false,
            FileTypeFilter =
            [
                new FilePickerFileType("圖片")
                {
                    Patterns = ["*.png", "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.webp"],
                    MimeTypes = ["image/*"],
                },
                AllFileType(),
            ],
        });
        if (files.Count == 0)
            return;

        var path = files[0].TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            StatusText = "無法讀取圖片路徑。";
            return;
        }

        try
        {
            PushUndo();
            var image = ImageBlock.FromFile(path);
            if (image.Preview is null)
            {
                StatusText = "無法解碼此圖片格式。";
                return;
            }

            var index = SelectedBlock is null ? Document.Blocks.Count : Document.Blocks.IndexOf(SelectedBlock) + 1;
            if (index < 0 || index > Document.Blocks.Count)
                index = Document.Blocks.Count;
            Document.Blocks.Insert(index, image);
            SelectedBlock = image;
            MarkDirty();
            StatusText = $"已插入圖片：{image.FileName}";
        }
        catch (Exception ex)
        {
            StatusText = $"插入圖片失敗：{ex.Message}";
        }
    }

    [RelayCommand]
    private void InsertHeading()
    {
        if (Document is null)
            return;
        PushUndo();
        var p = new ParagraphBlock
        {
            Text = "標題",
            Style = ParagraphStyleKind.Heading1,
            IsBold = true,
            FontSize = 28,
        };
        var index = SelectedBlock is null ? Document.Blocks.Count : Document.Blocks.IndexOf(SelectedBlock) + 1;
        if (index < 0 || index > Document.Blocks.Count)
            index = Document.Blocks.Count;
        Document.Blocks.Insert(index, p);
        SelectedBlock = p;
        MarkDirty();
        StatusText = "已插入標題";
    }

    [RelayCommand]
    private void DeleteBlock()
    {
        if (Document is null || SelectedBlock is null)
            return;
        if (Document.Blocks.Count <= 1)
        {
            StatusText = "至少需保留一個區塊。";
            return;
        }

        PushUndo();
        var index = Document.Blocks.IndexOf(SelectedBlock);
        Document.Blocks.Remove(SelectedBlock);
        SelectedBlock = Document.Blocks.ElementAtOrDefault(Math.Max(0, index - 1))
                        ?? Document.Blocks.FirstOrDefault();
        MarkDirty();
        StatusText = "已刪除區塊";
    }

    // ─── View / Find ────────────────────────────────────────

    [RelayCommand]
    private void ZoomIn()
    {
        Zoom = Math.Min(2.0, Math.Round(Zoom + 0.1, 2));
        StatusText = $"縮放 {ZoomText}";
    }

    [RelayCommand]
    private void ZoomOut()
    {
        Zoom = Math.Max(0.5, Math.Round(Zoom - 0.1, 2));
        StatusText = $"縮放 {ZoomText}";
    }

    [RelayCommand]
    private void ZoomReset()
    {
        Zoom = 1.0;
        StatusText = "縮放 100%";
    }

    [RelayCommand]
    private void ToggleFindPanel()
    {
        ShowFindPanel = !ShowFindPanel;
        if (!ShowFindPanel)
            FindMatchCount = 0;
    }

    [RelayCommand]
    private void FindNext()
    {
        if (Document is null || string.IsNullOrWhiteSpace(FindText))
        {
            StatusText = "請輸入搜尋文字。";
            return;
        }

        var paragraphs = Document.Blocks.OfType<ParagraphBlock>().ToList();
        var start = SelectedParagraph is null ? -1 : paragraphs.IndexOf(SelectedParagraph);
        for (var i = 1; i <= paragraphs.Count; i++)
        {
            var idx = (start + i) % paragraphs.Count;
            if (paragraphs[idx].Text.Contains(FindText, StringComparison.OrdinalIgnoreCase))
            {
                SelectedBlock = paragraphs[idx];
                StatusText = $"找到：第 {idx + 1} 段";
                return;
            }
        }

        StatusText = "找不到符合的內容。";
    }

    [RelayCommand]
    private void ReplaceOne()
    {
        if (SelectedParagraph is null || string.IsNullOrEmpty(FindText))
            return;
        if (!SelectedParagraph.Text.Contains(FindText, StringComparison.OrdinalIgnoreCase))
        {
            FindNext();
            return;
        }

        PushUndo();
        var idx = SelectedParagraph.Text.IndexOf(FindText, StringComparison.OrdinalIgnoreCase);
        SelectedParagraph.Text = SelectedParagraph.Text.Remove(idx, FindText.Length)
            .Insert(idx, ReplaceText ?? string.Empty);
        MarkDirty();
        StatusText = "已取代一處";
    }

    [RelayCommand]
    private void ReplaceAll()
    {
        if (Document is null || string.IsNullOrEmpty(FindText))
            return;
        PushUndo();
        var count = 0;
        foreach (var p in Document.Blocks.OfType<ParagraphBlock>())
        {
            if (string.IsNullOrEmpty(p.Text))
                continue;
            var before = p.Text;
            p.Text = ReplaceIgnoreCase(p.Text, FindText, ReplaceText ?? string.Empty);
            if (!string.Equals(before, p.Text, StringComparison.Ordinal))
                count++;
        }

        MarkDirty();
        StatusText = $"已取代 {count} 個段落中的內容";
    }

    private static string ReplaceIgnoreCase(string input, string search, string replacement)
    {
        var result = input;
        var index = 0;
        while (index < result.Length)
        {
            var found = result.IndexOf(search, index, StringComparison.OrdinalIgnoreCase);
            if (found < 0)
                break;
            result = result.Remove(found, search.Length).Insert(found, replacement);
            index = found + replacement.Length;
        }

        return result;
    }

    [RelayCommand]
    private void RefreshWordCount() => RefreshStatistics();

    // ─── Helpers ────────────────────────────────────────────

    private static Window? GetMainWindow()
    {
        if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            return desktop.MainWindow;
        return null;
    }

    private static IClipboard? GetClipboard()
    {
        var window = GetMainWindow();
        return window?.Clipboard;
    }

    private static FilePickerFileType DocxFileType() => new("Word 文件")
    {
        Patterns = ["*.docx"],
        AppleUniformTypeIdentifiers = ["org.openxmlformats.wordprocessingml.document"],
        MimeTypes = ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
    };

    private static FilePickerFileType AllFileType() => new("所有檔案") { Patterns = ["*.*"] };

    private sealed class DocDto
    {
        public string? Title { get; set; }
        public List<BlockDto>? Blocks { get; set; }
    }

    private sealed class BlockDto
    {
        public string Type { get; set; } = "p";
        public string? Text { get; set; }
        public string? Style { get; set; }
        public string? Align { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public double FontSize { get; set; }
        public string? List { get; set; }
        public List<List<string>>? Rows { get; set; }
        public string? ContentType { get; set; }
        public string? ImageBase64 { get; set; }
        public double DisplayWidth { get; set; }
        public double DisplayHeight { get; set; }
    }
}
