using System.IO;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxAvalonia.Services;

/// <summary>
/// Parses a .docx package with Open XML and builds Avalonia controls for on-screen preview.
/// </summary>
public sealed class DocxPreviewRenderer
{
    private static readonly SolidColorBrush HyperlinkBrush = new(Color.Parse("#0563C1"));
    private static readonly SolidColorBrush DefaultTextBrush = new(Color.Parse("#242424"));
    private static readonly SolidColorBrush TableHeaderBrush = new(Color.Parse("#F0F4F8"));
    private static readonly SolidColorBrush TableBorderBrush = new(Color.Parse("#CCCCCC"));

    private MainDocumentPart? _mainPart;
    private readonly Dictionary<string, NumberingLevelData> _numbering = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _listCounters = new(StringComparer.Ordinal);

    public Control Render(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("找不到指定的文件。", filePath);

        if (!string.Equals(Path.GetExtension(filePath), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("僅支援 .docx 格式（Office Open XML）。");

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Render(stream);
    }

    public Control Render(Stream stream)
    {
        _numbering.Clear();
        _listCounters.Clear();

        using var document = WordprocessingDocument.Open(stream, false);
        _mainPart = document.MainDocumentPart
            ?? throw new InvalidOperationException("無效的 Word 文件：缺少主文件部件。");

        LoadNumbering();

        var body = _mainPart.Document?.Body
            ?? throw new InvalidOperationException("無效的 Word 文件：缺少文件本文。");

        var pageContent = new StackPanel
        {
            Orientation = Orientation.Vertical,
            Spacing = 0,
        };

        foreach (var element in body.Elements())
        {
            switch (element)
            {
                case W.Paragraph paragraph:
                    var paragraphControl = BuildParagraph(paragraph);
                    if (paragraphControl is not null)
                        pageContent.Children.Add(paragraphControl);
                    break;

                case W.Table table:
                    pageContent.Children.Add(BuildTable(table));
                    break;
            }
        }

        if (pageContent.Children.Count == 0)
        {
            pageContent.Children.Add(new TextBlock
            {
                Classes = { "docx-empty" },
                Text = "此文件沒有可預覽的內容。",
            });
        }

        return new Border
        {
            Classes = { "docx-page" },
            Child = pageContent,
        };
    }

    private void LoadNumbering()
    {
        var numberingPart = _mainPart?.NumberingDefinitionsPart;
        if (numberingPart?.Numbering is null)
            return;

        var abstractNums = numberingPart.Numbering
            .Elements<W.AbstractNum>()
            .ToDictionary(a => a.AbstractNumberId?.Value ?? -1);

        foreach (var num in numberingPart.Numbering.Elements<W.NumberingInstance>())
        {
            var numIdValue = num.NumberID?.Value;
            if (numIdValue is null)
                continue;

            var numId = numIdValue.Value.ToString();
            var abstractId = num.AbstractNumId?.Val?.Value ?? -1;
            if (!abstractNums.TryGetValue(abstractId, out var abstractNum))
                continue;

            var level0 = abstractNum.Elements<W.Level>().FirstOrDefault(l => l.LevelIndex?.Value == 0);
            if (level0 is null)
                continue;

            var format = level0.NumberingFormat?.Val?.Value;
            var isBullet = format == W.NumberFormatValues.Bullet;
            var levelText = level0.LevelText?.Val?.Value ?? (isBullet ? "•" : "%1.");

            _numbering[numId] = new NumberingLevelData(isBullet, levelText);
        }
    }

    private Control? BuildParagraph(W.Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var classes = ResolveParagraphClasses(styleId);
        var alignment = ResolveAlignment(paragraph.ParagraphProperties?.Justification?.Val?.Value);

        var images = ExtractImages(paragraph).ToList();
        var inlines = BuildInlines(paragraph).ToList();

        if (images.Count == 0 && inlines.Count == 0)
        {
            return new TextBlock
            {
                Classes = { classes },
                Text = " ",
                Margin = new Thickness(0, 0, 0, 6),
            };
        }

        var container = new StackPanel { Orientation = Orientation.Vertical };

        if (inlines.Count > 0)
        {
            var textBlock = new TextBlock
            {
                Classes = { classes },
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = alignment,
            };

            var listPrefix = TryGetListPrefix(paragraph);
            if (listPrefix is not null)
            {
                textBlock.Classes.Clear();
                textBlock.Classes.Add("docx-list-item");
                textBlock.Inlines!.Add(new Run(listPrefix)
                {
                    FontWeight = FontWeight.Normal,
                    Foreground = DefaultTextBrush,
                });
            }

            foreach (var inline in inlines)
                textBlock.Inlines!.Add(inline);

            container.Children.Add(textBlock);
        }

        foreach (var image in images)
            container.Children.Add(image);

        // Always return a control that has no parent yet.
        // Returning Children[0] while it is still parented to `container` causes:
        // "already has a visual parent StackPanel while trying to add it as a child of StackPanel"
        if (container.Children.Count == 1)
        {
            var only = container.Children[0];
            container.Children.Clear();
            return only;
        }

        return container;
    }

    private static string ResolveParagraphClasses(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId))
            return "docx-paragraph";

        var id = styleId.ToLowerInvariant();
        if (id is "heading1" or "title" or "標題1")
            return "docx-heading1";
        if (id is "heading2" or "subtitle" or "標題2")
            return "docx-heading2";
        if (id is "heading3" or "標題3")
            return "docx-heading3";

        return "docx-paragraph";
    }

    private static TextAlignment ResolveAlignment(EnumValue<W.JustificationValues>? value)
    {
        if (value is null)
            return TextAlignment.Left;

        if (value == W.JustificationValues.Center)
            return TextAlignment.Center;
        if (value == W.JustificationValues.Right)
            return TextAlignment.Right;
        if (value == W.JustificationValues.Both)
            return TextAlignment.Justify;

        return TextAlignment.Left;
    }

    private string? TryGetListPrefix(W.Paragraph paragraph)
    {
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        var numIdVal = numPr?.NumberingId?.Val?.Value;
        if (numIdVal is null)
            return null;

        var numId = numIdVal.Value.ToString();
        if (!_numbering.TryGetValue(numId, out var data))
            return "• ";

        if (data.IsBullet)
            return "• ";

        if (!_listCounters.TryGetValue(numId, out var counter))
            counter = 0;

        counter++;
        _listCounters[numId] = counter;

        var text = data.LevelText.Replace("%1", counter.ToString(), StringComparison.Ordinal);
        if (text.EndsWith('.'))
            text += " ";
        else if (!text.EndsWith(' '))
            text += " ";

        return text;
    }

    private IEnumerable<Inline> BuildInlines(W.Paragraph paragraph)
    {
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case W.Run run:
                    foreach (var inline in BuildRunInlines(run, isHyperlink: false))
                        yield return inline;
                    break;

                case W.Hyperlink hyperlink:
                    foreach (var inline in BuildHyperlinkInlines(hyperlink))
                        yield return inline;
                    break;
            }
        }
    }

    private IEnumerable<Inline> BuildRunInlines(W.Run run, bool isHyperlink)
    {
        if (run.Descendants<W.Drawing>().Any() && !run.Descendants<W.Text>().Any())
            yield break;

        var props = run.RunProperties;
        var isBold = props?.Bold is not null && (props.Bold.Val is null || props.Bold.Val.Value);
        var isItalic = props?.Italic is not null && (props.Italic.Val is null || props.Italic.Val.Value);
        var isUnderline = props?.Underline is not null
            && props.Underline.Val is not null
            && props.Underline.Val.Value != W.UnderlineValues.None;

        if (isHyperlink)
            isUnderline = true;

        var fontSize = ResolveFontSize(props);
        var foreground = isHyperlink
            ? HyperlinkBrush
            : ResolveColor(props) ?? DefaultTextBrush;
        var fontFamily = ResolveFontFamily(props);

        foreach (var element in run.ChildElements)
        {
            switch (element)
            {
                case W.Text text:
                    yield return CreateRun(text.Text ?? string.Empty, isBold, isItalic, isUnderline, fontSize, foreground, fontFamily);
                    break;

                case W.Break br when br.Type?.Value == W.BreakValues.Page:
                    yield return new LineBreak();
                    yield return new LineBreak();
                    break;

                case W.Break:
                    yield return new LineBreak();
                    break;

                case W.TabChar:
                    yield return CreateRun("    ", isBold, isItalic, isUnderline, fontSize, foreground, fontFamily);
                    break;

                case W.SymbolChar:
                    yield return CreateRun("•", isBold, isItalic, isUnderline, fontSize, foreground, fontFamily);
                    break;
            }
        }
    }

    private IEnumerable<Inline> BuildHyperlinkInlines(W.Hyperlink hyperlink)
    {
        foreach (var run in hyperlink.Elements<W.Run>())
        {
            foreach (var inline in BuildRunInlines(run, isHyperlink: true))
                yield return inline;
        }
    }

    private static Run CreateRun(
        string text,
        bool bold,
        bool italic,
        bool underline,
        double? fontSize,
        IBrush foreground,
        FontFamily? fontFamily)
    {
        var run = new Run(text)
        {
            FontWeight = bold ? FontWeight.Bold : FontWeight.Normal,
            FontStyle = italic ? FontStyle.Italic : FontStyle.Normal,
            Foreground = foreground,
        };

        if (underline)
            run.TextDecorations = TextDecorations.Underline;

        if (fontSize is not null)
            run.FontSize = fontSize.Value;

        if (fontFamily is not null)
            run.FontFamily = fontFamily;

        return run;
    }

    private static double? ResolveFontSize(W.RunProperties? props)
    {
        var halfPoints = props?.FontSize?.Val?.Value;
        if (halfPoints is null)
            return null;

        if (double.TryParse(halfPoints, out var value))
            return value / 2.0;

        return null;
    }

    private static IBrush? ResolveColor(W.RunProperties? props)
    {
        var colorVal = props?.Color?.Val?.Value;
        if (string.IsNullOrWhiteSpace(colorVal) || colorVal.Equals("auto", StringComparison.OrdinalIgnoreCase))
            return null;

        try
        {
            if (colorVal.Length == 6)
                return new SolidColorBrush(Color.Parse("#" + colorVal));
            if (colorVal.StartsWith('#'))
                return new SolidColorBrush(Color.Parse(colorVal));
        }
        catch
        {
            // Fall through.
        }

        return null;
    }

    private static FontFamily? ResolveFontFamily(W.RunProperties? props)
    {
        var name = props?.RunFonts?.Ascii?.Value
            ?? props?.RunFonts?.HighAnsi?.Value
            ?? props?.RunFonts?.EastAsia?.Value;

        if (string.IsNullOrWhiteSpace(name))
            return null;

        try
        {
            return new FontFamily(name);
        }
        catch
        {
            return null;
        }
    }

    private IEnumerable<Control> ExtractImages(W.Paragraph paragraph)
    {
        if (_mainPart is null)
            yield break;

        foreach (var drawing in paragraph.Descendants<W.Drawing>())
        {
            var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
            var embedId = blip?.Embed?.Value;
            if (string.IsNullOrEmpty(embedId))
                continue;

            OpenXmlPart? part;
            try
            {
                part = _mainPart.GetPartById(embedId);
            }
            catch
            {
                continue;
            }

            if (part is not ImagePart imagePart)
                continue;

            Bitmap bitmap;
            try
            {
                using var imageStream = imagePart.GetStream();
                using var ms = new MemoryStream();
                imageStream.CopyTo(ms);
                ms.Position = 0;
                bitmap = new Bitmap(ms);
            }
            catch
            {
                continue;
            }

            double? maxWidth = null;
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            if (extent?.Cx?.Value is { } cxFromExtent)
            {
                maxWidth = cxFromExtent / 914400.0 * 96.0;
            }
            else
            {
                var pic = drawing.Descendants<PIC.Picture>().FirstOrDefault();
                var cx = pic?.ShapeProperties?.Transform2D?.Extents?.Cx?.Value;
                if (cx is not null)
                    maxWidth = cx.Value / 914400.0 * 96.0;
            }

            var image = new Image
            {
                Classes = { "docx-image" },
                Source = bitmap,
            };

            if (maxWidth is > 0)
                image.MaxWidth = Math.Min(maxWidth.Value, 640);

            yield return image;
        }
    }

    private Control BuildTable(W.Table table)
    {
        var grid = new Grid
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };

        var rows = table.Elements<W.TableRow>().ToList();
        if (rows.Count == 0)
            return new Border { Classes = { "docx-table" }, Child = grid };

        var columnCount = rows.Max(r => r.Elements<W.TableCell>().Count());
        for (var c = 0; c < columnCount; c++)
            grid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));

        for (var r = 0; r < rows.Count; r++)
            grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        for (var r = 0; r < rows.Count; r++)
        {
            var cells = rows[r].Elements<W.TableCell>().ToList();
            var isHeader = r == 0;

            for (var c = 0; c < cells.Count; c++)
            {
                var cellPanel = new StackPanel { Orientation = Orientation.Vertical };

                foreach (var paragraph in cells[c].Elements<W.Paragraph>())
                {
                    var p = BuildParagraph(paragraph);
                    if (p is not null)
                        cellPanel.Children.Add(p);
                }

                if (cellPanel.Children.Count == 0)
                    cellPanel.Children.Add(new TextBlock { Text = " " });

                var cellBorder = new Border
                {
                    Classes = { isHeader ? "docx-table-header-cell" : "docx-table-cell" },
                    Background = isHeader ? TableHeaderBrush : Brushes.White,
                    BorderBrush = TableBorderBrush,
                    BorderThickness = new Thickness(0.5),
                    Padding = new Thickness(8, 6),
                    Child = cellPanel,
                };

                Grid.SetRow(cellBorder, r);
                Grid.SetColumn(cellBorder, c);
                grid.Children.Add(cellBorder);
            }
        }

        return new Border
        {
            Classes = { "docx-table" },
            BorderBrush = TableBorderBrush,
            BorderThickness = new Thickness(1),
            Child = grid,
            Margin = new Thickness(0, 8, 0, 16),
        };
    }

    private sealed record NumberingLevelData(bool IsBullet, string LevelText);
}
