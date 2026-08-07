using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;

namespace DocxAvalonia.Models;

public enum ParagraphStyleKind
{
    Normal,
    Heading1,
    Heading2,
    Heading3,
}

public enum ParagraphAlignmentKind
{
    Left,
    Center,
    Right,
    Justify,
}

public enum ListKind
{
    None,
    Bullet,
    Numbered,
}

/// <summary>Editable document root.</summary>
public partial class WordDocument : ObservableObject
{
    public ObservableCollection<DocumentBlock> Blocks { get; } = new();

    [ObservableProperty]
    private string _title = "未命名文件";

    public static WordDocument CreateBlank()
    {
        var doc = new WordDocument();
        doc.Blocks.Add(new ParagraphBlock { Text = string.Empty });
        return doc;
    }

    public string GetPlainText()
    {
        var sb = new StringBuilder();
        foreach (var block in Blocks)
        {
            switch (block)
            {
                case ParagraphBlock p:
                    sb.AppendLine(p.Text);
                    break;
                case TableBlock t:
                    foreach (var row in t.Rows)
                        sb.AppendLine(string.Join('\t', row.Cells.Select(c => c.Text)));
                    break;
            }
        }

        return sb.ToString();
    }

    public (int Words, int Chars, int CharsNoSpaces, int Paragraphs) GetStatistics()
    {
        var text = GetPlainText();
        var chars = text.Length;
        var charsNoSpaces = text.Count(c => !char.IsWhiteSpace(c));
        var words = text
            .Split([' ', '\t', '\r', '\n', '　'], StringSplitOptions.RemoveEmptyEntries)
            .Length;
        var paragraphs = Blocks.OfType<ParagraphBlock>().Count(p => !string.IsNullOrWhiteSpace(p.Text));
        return (words, chars, charsNoSpaces, paragraphs);
    }
}

public abstract partial class DocumentBlock : ObservableObject
{
    public string Id { get; } = Guid.NewGuid().ToString("N");
}

public partial class ParagraphBlock : DocumentBlock
{
    [ObservableProperty]
    private string _text = string.Empty;

    [ObservableProperty]
    private ParagraphStyleKind _style = ParagraphStyleKind.Normal;

    [ObservableProperty]
    private ParagraphAlignmentKind _alignment = ParagraphAlignmentKind.Left;

    [ObservableProperty]
    private bool _isBold;

    [ObservableProperty]
    private bool _isItalic;

    [ObservableProperty]
    private bool _isUnderline;

    [ObservableProperty]
    private double _fontSize = 14;

    /// <summary>UI / OpenXML font family name (e.g. Microsoft JhengHei, Calibri).</summary>
    [ObservableProperty]
    private string _fontFamily = "Microsoft JhengHei";

    [ObservableProperty]
    private ListKind _listKind = ListKind.None;

    partial void OnStyleChanged(ParagraphStyleKind value)
    {
        // Apply typical heading sizes when style changes.
        FontSize = value switch
        {
            ParagraphStyleKind.Heading1 => 28,
            ParagraphStyleKind.Heading2 => 22,
            ParagraphStyleKind.Heading3 => 18,
            _ => FontSize < 12 ? 14 : FontSize,
        };
        if (value is ParagraphStyleKind.Heading1 or ParagraphStyleKind.Heading2 or ParagraphStyleKind.Heading3)
            IsBold = true;
    }
}

public partial class TableBlock : DocumentBlock
{
    public const int MaxRows = 100;
    public const int MaxColumns = 20;

    public ObservableCollection<TableRowBlock> Rows { get; } = new();

    public int ColumnCount =>
        Rows.Count == 0 ? 0 : Rows.Max(r => r.Cells.Count);

    public int RowCount => Rows.Count;

    public TableBlock()
    {
        Rows.CollectionChanged += (_, _) => NotifyShapeChanged();
    }

    public static TableBlock Create(int rows, int cols)
    {
        rows = Math.Clamp(rows, 1, MaxRows);
        cols = Math.Clamp(cols, 1, MaxColumns);
        var table = new TableBlock();
        for (var r = 0; r < rows; r++)
        {
            var row = new TableRowBlock();
            for (var c = 0; c < cols; c++)
                row.Cells.Add(new TableCellBlock { Text = r == 0 ? $"欄{c + 1}" : string.Empty });
            table.Rows.Add(row);
        }

        return table;
    }

    /// <summary>Pad all rows so each has the same column count.</summary>
    public void EnsureRectangular()
    {
        var cols = Math.Max(1, ColumnCount);
        foreach (var row in Rows)
        {
            while (row.Cells.Count < cols)
                row.Cells.Add(new TableCellBlock());
        }
    }

    /// <summary>Insert a row after <paramref name="afterRowIndex"/> (-1 = before first; null = append).</summary>
    public TableRowBlock InsertRow(int? afterRowIndex = null)
    {
        EnsureRectangular();
        if (Rows.Count >= MaxRows)
            throw new InvalidOperationException($"表格列數上限為 {MaxRows}。");

        var cols = Math.Max(1, ColumnCount);
        var row = new TableRowBlock();
        for (var c = 0; c < cols; c++)
            row.Cells.Add(new TableCellBlock());

        if (afterRowIndex is null)
            Rows.Add(row);
        else if (afterRowIndex < 0)
            Rows.Insert(0, row);
        else if (afterRowIndex >= Rows.Count - 1)
            Rows.Add(row);
        else
            Rows.Insert(afterRowIndex.Value + 1, row);

        NotifyShapeChanged();
        return row;
    }

    /// <summary>Insert a column after <paramref name="afterColIndex"/> (-1 = before first; null = append).</summary>
    public void InsertColumn(int? afterColIndex = null)
    {
        EnsureRectangular();
        var cols = ColumnCount;
        if (cols >= MaxColumns)
            throw new InvalidOperationException($"表格欄數上限為 {MaxColumns}。");

        if (Rows.Count == 0)
        {
            var row = new TableRowBlock();
            row.Cells.Add(new TableCellBlock());
            Rows.Add(row);
            NotifyShapeChanged();
            return;
        }

        foreach (var row in Rows)
        {
            var cell = new TableCellBlock();
            if (afterColIndex is null)
                row.Cells.Add(cell);
            else if (afterColIndex < 0)
                row.Cells.Insert(0, cell);
            else if (afterColIndex >= row.Cells.Count - 1)
                row.Cells.Add(cell);
            else
                row.Cells.Insert(afterColIndex.Value + 1, cell);
        }

        NotifyShapeChanged();
    }

    public bool DeleteRowAt(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= Rows.Count)
            return false;
        if (Rows.Count <= 1)
            throw new InvalidOperationException("至少需保留一列。");
        Rows.RemoveAt(rowIndex);
        NotifyShapeChanged();
        return true;
    }

    public bool DeleteColumnAt(int colIndex)
    {
        EnsureRectangular();
        var cols = ColumnCount;
        if (colIndex < 0 || colIndex >= cols)
            return false;
        if (cols <= 1)
            throw new InvalidOperationException("至少需保留一欄。");

        foreach (var row in Rows)
        {
            if (colIndex < row.Cells.Count)
                row.Cells.RemoveAt(colIndex);
        }

        NotifyShapeChanged();
        return true;
    }

    public bool TryFindCell(TableCellBlock cell, out int rowIndex, out int colIndex)
    {
        for (var r = 0; r < Rows.Count; r++)
        {
            var c = Rows[r].Cells.IndexOf(cell);
            if (c >= 0)
            {
                rowIndex = r;
                colIndex = c;
                return true;
            }
        }

        rowIndex = -1;
        colIndex = -1;
        return false;
    }

    private void NotifyShapeChanged()
    {
        OnPropertyChanged(nameof(RowCount));
        OnPropertyChanged(nameof(ColumnCount));
        foreach (var row in Rows)
            row.NotifyColumnCountChanged();
    }
}

public partial class TableRowBlock : ObservableObject
{
    public ObservableCollection<TableCellBlock> Cells { get; } = new();

    public int ColumnCount => Cells.Count;

    public TableRowBlock()
    {
        Cells.CollectionChanged += (_, _) => NotifyColumnCountChanged();
    }

    public void NotifyColumnCountChanged() => OnPropertyChanged(nameof(ColumnCount));
}

public partial class TableCellBlock : ObservableObject
{
    [ObservableProperty]
    private string _text = string.Empty;

    [ObservableProperty]
    private double _fontSize = 14;

    [ObservableProperty]
    private string _fontFamily = "Microsoft JhengHei";

    [ObservableProperty]
    private bool _isBold;

    [ObservableProperty]
    private bool _isItalic;

    [ObservableProperty]
    private bool _isUnderline;

    [ObservableProperty]
    private ParagraphAlignmentKind _alignment = ParagraphAlignmentKind.Left;
}

/// <summary>Embedded image block loaded from or saved to .docx.</summary>
public partial class ImageBlock : DocumentBlock
{
    private readonly object _previewLock = new();
    private Bitmap? _preview;
    private bool _previewAttempted;
    private byte[] _imageBytes = Array.Empty<byte>();

    public byte[] ImageBytes
    {
        get => _imageBytes;
        set
        {
            // Do not Dispose previous Bitmap here — may still be bound to UI.
            if (SetProperty(ref _imageBytes, value ?? Array.Empty<byte>()))
            {
                lock (_previewLock)
                {
                    _preview = null;
                    _previewAttempted = false;
                }

                OnPropertyChanged(nameof(Preview));
                OnPropertyChanged(nameof(HasPreview));
                OnPropertyChanged(nameof(IsDecodeFailed));
            }
        }
    }

    /// <summary>MIME / OpenXML content type, e.g. image/png, image/jpeg.</summary>
    [ObservableProperty]
    private string _contentType = "image/png";

    [ObservableProperty]
    private string? _fileName;

    /// <summary>Display width in DIPs (UI pixels).</summary>
    [ObservableProperty]
    private double _displayWidth = 400;

    /// <summary>Display height in DIPs; 0 = keep aspect ratio from width.</summary>
    [ObservableProperty]
    private double _displayHeight;

    /// <summary>True when this image is the active selection (shows resize handles).</summary>
    [ObservableProperty]
    private bool _isSelected;

    /// <summary>Natural pixel size once decoded (0 if unknown).</summary>
    public int PixelWidth { get; private set; }

    public int PixelHeight { get; private set; }

    public const double MinDisplayWidth = 40;
    public const double MaxDisplayWidth = 720;
    public const double MinDisplayHeight = 24;
    public const double MaxDisplayHeight = 2000;

    /// <summary>Effective display height (never zero once size is known).</summary>
    public double ResolvedHeight
    {
        get
        {
            if (DisplayHeight > 0 && !double.IsNaN(DisplayHeight))
                return DisplayHeight;
            if (PixelWidth > 0 && PixelHeight > 0 && DisplayWidth > 0)
                return DisplayWidth * PixelHeight / PixelWidth;
            return Math.Max(MinDisplayHeight, DisplayWidth * 0.75);
        }
    }

    /// <summary>Resize keeping aspect ratio when pixel size is known.</summary>
    public void ResizeToWidth(double width)
    {
        if (double.IsNaN(width) || double.IsInfinity(width))
            return;
        width = Math.Clamp(width, MinDisplayWidth, MaxDisplayWidth);
        DisplayWidth = width;
        if (PixelWidth > 0 && PixelHeight > 0)
            DisplayHeight = width * PixelHeight / PixelWidth;
        else if (DisplayHeight <= 0 || double.IsNaN(DisplayHeight))
            DisplayHeight = width * 0.75;
    }

    /// <summary>Free resize (may change aspect ratio). Used by 8-direction handles.</summary>
    public void ResizeToSize(double width, double height)
    {
        if (double.IsNaN(width) || double.IsInfinity(width)
            || double.IsNaN(height) || double.IsInfinity(height))
            return;

        DisplayWidth = Math.Clamp(width, MinDisplayWidth, MaxDisplayWidth);
        DisplayHeight = Math.Clamp(height, MinDisplayHeight, MaxDisplayHeight);
    }

    /// <summary>
    /// Apply drag delta from an 8-direction handle.
    /// Handles: n, s, e, w, ne, nw, se, sw.
    /// </summary>
    public void ApplyHandleDelta(string handle, double dx, double dy, double startWidth, double startHeight)
    {
        var w = startWidth;
        var h = startHeight;
        handle = (handle ?? string.Empty).Trim().ToLowerInvariant();

        // Horizontal: east grows, west shrinks (inline images only change size, not origin).
        if (handle is "e" or "ne" or "se")
            w = startWidth + dx;
        else if (handle is "w" or "nw" or "sw")
            w = startWidth - dx;

        // Vertical: south grows, north shrinks.
        if (handle is "s" or "se" or "sw")
            h = startHeight + dy;
        else if (handle is "n" or "ne" or "nw")
            h = startHeight - dy;

        ResizeToSize(w, h);
    }

    /// <summary>Scale current display size by a factor (e.g. 1.15 = larger).</summary>
    public void ResizeByFactor(double factor)
    {
        if (factor <= 0 || double.IsNaN(factor) || double.IsInfinity(factor))
            return;
        var w = DisplayWidth > 0 ? DisplayWidth : 400;
        var h = ResolvedHeight;
        ResizeToSize(w * factor, h * factor);
    }

    public string SizeLabel
    {
        get
        {
            var w = (int)Math.Round(DisplayWidth);
            var h = (int)Math.Round(ResolvedHeight);
            return $"{w} × {h}";
        }
    }

    partial void OnDisplayWidthChanged(double value)
    {
        OnPropertyChanged(nameof(SizeLabel));
        OnPropertyChanged(nameof(ResolvedHeight));
    }

    partial void OnDisplayHeightChanged(double value)
    {
        OnPropertyChanged(nameof(SizeLabel));
        OnPropertyChanged(nameof(ResolvedHeight));
    }

    public bool HasPreview => Preview is not null;

    public bool IsDecodeFailed
    {
        get
        {
            _ = Preview; // ensure attempt
            return _previewAttempted && _preview is null && _imageBytes.Length > 0;
        }
    }

    public Bitmap? Preview
    {
        get
        {
            lock (_previewLock)
            {
                if (_preview is not null || _previewAttempted)
                    return _preview;

                _previewAttempted = true;
                TryCreatePreviewUnlocked();
                return _preview;
            }
        }
    }

    /// <summary>Decode bitmap on the UI thread after document load.</summary>
    public void EnsurePreview()
    {
        lock (_previewLock)
        {
            if (_preview is not null || _previewAttempted)
                return;
            _previewAttempted = true;
            TryCreatePreviewUnlocked();
        }

        OnPropertyChanged(nameof(Preview));
        OnPropertyChanged(nameof(HasPreview));
        OnPropertyChanged(nameof(IsDecodeFailed));
    }

    private void TryCreatePreviewUnlocked()
    {
        if (_imageBytes.Length == 0)
            return;

        if (!IsDecodableContentType(ContentType) && !LooksLikeRasterImage(_imageBytes))
            return;

        try
        {
            // Copy so Avalonia/Skia owns independent memory; never dispose stream while bitmap lives.
            var copy = new byte[_imageBytes.Length];
            Buffer.BlockCopy(_imageBytes, 0, copy, 0, _imageBytes.Length);
            using var ms = new MemoryStream(copy, writable: false);
            var bmp = new Bitmap(ms);
            _preview = bmp;
            PixelWidth = bmp.PixelSize.Width;
            PixelHeight = bmp.PixelSize.Height;

            if (DisplayWidth <= 0 || double.IsNaN(DisplayWidth) || double.IsInfinity(DisplayWidth))
                DisplayWidth = 400;
            DisplayWidth = Math.Clamp(DisplayWidth, 40, 720);

            if ((DisplayHeight <= 0 || double.IsNaN(DisplayHeight)) && PixelWidth > 0)
                DisplayHeight = DisplayWidth * PixelHeight / PixelWidth;
        }
        catch
        {
            _preview = null;
        }
    }

    public static bool IsDecodableContentType(string? contentType)
    {
        if (string.IsNullOrWhiteSpace(contentType))
            return true; // try by magic bytes
        var ct = contentType.ToLowerInvariant();
        // EMF/WMF/SVG often crash or are unsupported by Skia bitmap loader — skip decode.
        if (ct.Contains("emf") || ct.Contains("wmf") || ct.Contains("svg") || ct.Contains("x-emf") || ct.Contains("x-wmf"))
            return false;
        return ct.Contains("png")
               || ct.Contains("jpeg")
               || ct.Contains("jpg")
               || ct.Contains("gif")
               || ct.Contains("bmp")
               || ct.Contains("webp")
               || ct.Contains("tiff")
               || ct.Contains("image/");
    }

    private static bool LooksLikeRasterImage(byte[] bytes)
    {
        if (bytes.Length < 4)
            return false;
        // PNG
        if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
            return true;
        // JPEG
        if (bytes[0] == 0xFF && bytes[1] == 0xD8)
            return true;
        // GIF
        if (bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46)
            return true;
        // BMP
        if (bytes[0] == 0x42 && bytes[1] == 0x4D)
            return true;
        // WEBP (RIFF....WEBP)
        if (bytes.Length >= 12
            && bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46
            && bytes[8] == 0x57 && bytes[9] == 0x45 && bytes[10] == 0x42 && bytes[11] == 0x50)
            return true;
        return false;
    }

    public static ImageBlock FromFile(string path, double maxWidth = 640)
    {
        var bytes = File.ReadAllBytes(path);
        var ext = Path.GetExtension(path).ToLowerInvariant();
        var contentType = ext switch
        {
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            ".webp" => "image/webp",
            _ => "image/png",
        };

        var block = new ImageBlock
        {
            ImageBytes = bytes,
            ContentType = contentType,
            FileName = Path.GetFileName(path),
            DisplayWidth = maxWidth,
        };
        block.EnsurePreview();

        if (block.PixelWidth > 0)
        {
            block.DisplayWidth = Math.Min(maxWidth, block.PixelWidth);
            block.DisplayHeight = block.DisplayWidth * block.PixelHeight / block.PixelWidth;
        }

        return block;
    }
}

