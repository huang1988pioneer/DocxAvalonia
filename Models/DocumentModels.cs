using System.Collections.ObjectModel;
using System.Text;
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
    public ObservableCollection<TableRowBlock> Rows { get; } = new();

    public static TableBlock Create(int rows, int cols)
    {
        rows = Math.Clamp(rows, 1, 20);
        cols = Math.Clamp(cols, 1, 10);
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
}

public partial class TableRowBlock : ObservableObject
{
    public ObservableCollection<TableCellBlock> Cells { get; } = new();
}

public partial class TableCellBlock : ObservableObject
{
    [ObservableProperty]
    private string _text = string.Empty;
}
