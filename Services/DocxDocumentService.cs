using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxAvalonia.Models;

namespace DocxAvalonia.Services;

/// <summary>Load / save editable document models as .docx.</summary>
public sealed class DocxDocumentService
{
    public WordDocument Load(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("找不到指定的文件。", path);

        if (!string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("僅支援 .docx 格式。");

        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Load(stream, Path.GetFileNameWithoutExtension(path));
    }

    public WordDocument Load(Stream stream, string? title = null)
    {
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart
            ?? throw new InvalidOperationException("無效的 Word 文件：缺少主文件部件。");
        var body = main.Document?.Body
            ?? throw new InvalidOperationException("無效的 Word 文件：缺少文件本文。");

        var numbering = LoadNumberingMap(main);
        var doc = new WordDocument { Title = title ?? "文件" };

        foreach (var element in body.Elements())
        {
            switch (element)
            {
                case Paragraph paragraph:
                    doc.Blocks.Add(ConvertParagraph(paragraph, numbering));
                    break;
                case Table table:
                    doc.Blocks.Add(ConvertTable(table));
                    break;
            }
        }

        if (doc.Blocks.Count == 0)
            doc.Blocks.Add(new ParagraphBlock());

        return doc;
    }

    public void Save(WordDocument model, string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        if (File.Exists(path))
            File.Delete(path);

        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body());
        EnsureStyles(main);
        EnsureNumbering(main);

        var body = main.Document.Body!;
        var bulletCount = 0;
        var numberCount = 0;

        foreach (var block in model.Blocks)
        {
            switch (block)
            {
                case ParagraphBlock p:
                    body.AppendChild(CreateParagraph(p, ref bulletCount, ref numberCount));
                    break;
                case TableBlock t:
                    body.AppendChild(CreateTable(t));
                    break;
            }
        }

        if (!body.Elements().Any())
            body.AppendChild(new Paragraph(new Run(new Text(string.Empty))));

        main.Document.Save();
    }

    private static Dictionary<int, ListKind> LoadNumberingMap(MainDocumentPart main)
    {
        var map = new Dictionary<int, ListKind>();
        var numberingPart = main.NumberingDefinitionsPart;
        if (numberingPart?.Numbering is null)
            return map;

        var abstracts = numberingPart.Numbering.Elements<AbstractNum>()
            .ToDictionary(a => a.AbstractNumberId?.Value ?? -1);

        foreach (var num in numberingPart.Numbering.Elements<NumberingInstance>())
        {
            var id = num.NumberID?.Value;
            if (id is null)
                continue;
            var absId = num.AbstractNumId?.Val?.Value ?? -1;
            if (!abstracts.TryGetValue(absId, out var abs))
                continue;
            var level = abs.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == 0);
            var format = level?.NumberingFormat?.Val?.Value;
            map[id.Value] = format == NumberFormatValues.Bullet ? ListKind.Bullet : ListKind.Numbered;
        }

        return map;
    }

    private static ParagraphBlock ConvertParagraph(Paragraph paragraph, Dictionary<int, ListKind> numbering)
    {
        var sb = new StringBuilder();
        var bold = false;
        var italic = false;
        var underline = false;
        double? fontSize = null;
        var runCount = 0;
        var boldCount = 0;
        var italicCount = 0;
        var underlineCount = 0;

        foreach (var run in paragraph.Elements<Run>())
        {
            var text = string.Concat(run.Elements<Text>().Select(t => t.Text));
            if (string.IsNullOrEmpty(text) && run.Elements<Break>().Any())
            {
                sb.AppendLine();
                continue;
            }

            if (string.IsNullOrEmpty(text))
                continue;

            sb.Append(text);
            runCount++;
            var props = run.RunProperties;
            if (props?.Bold is not null && (props.Bold.Val is null || props.Bold.Val.Value))
                boldCount++;
            if (props?.Italic is not null && (props.Italic.Val is null || props.Italic.Val.Value))
                italicCount++;
            if (props?.Underline is not null
                && props.Underline.Val is not null
                && props.Underline.Val.Value != UnderlineValues.None)
                underlineCount++;

            var half = props?.FontSize?.Val?.Value;
            if (half is not null && double.TryParse(half, out var hp))
                fontSize ??= hp / 2.0;
        }

        // Also collect hyperlink runs
        foreach (var hyperlink in paragraph.Elements<Hyperlink>())
        {
            foreach (var run in hyperlink.Elements<Run>())
            {
                var text = string.Concat(run.Elements<Text>().Select(t => t.Text));
                if (!string.IsNullOrEmpty(text))
                    sb.Append(text);
            }
        }

        if (runCount > 0)
        {
            bold = boldCount * 2 >= runCount;
            italic = italicCount * 2 >= runCount;
            underline = underlineCount * 2 >= runCount;
        }

        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var style = styleId?.ToLowerInvariant() switch
        {
            "heading1" or "title" or "標題1" => ParagraphStyleKind.Heading1,
            "heading2" or "subtitle" or "標題2" => ParagraphStyleKind.Heading2,
            "heading3" or "標題3" => ParagraphStyleKind.Heading3,
            _ => ParagraphStyleKind.Normal,
        };

        var align = paragraph.ParagraphProperties?.Justification?.Val?.Value;
        var alignment = ParagraphAlignmentKind.Left;
        if (align == JustificationValues.Center)
            alignment = ParagraphAlignmentKind.Center;
        else if (align == JustificationValues.Right)
            alignment = ParagraphAlignmentKind.Right;
        else if (align == JustificationValues.Both)
            alignment = ParagraphAlignmentKind.Justify;

        var listKind = ListKind.None;
        var numId = paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
        if (numId is not null && numbering.TryGetValue(numId.Value, out var lk))
            listKind = lk;
        else if (numId is not null)
            listKind = ListKind.Bullet;

        return new ParagraphBlock
        {
            Text = sb.ToString(),
            Style = style,
            Alignment = alignment,
            IsBold = bold || style != ParagraphStyleKind.Normal,
            IsItalic = italic,
            IsUnderline = underline,
            FontSize = fontSize ?? style switch
            {
                ParagraphStyleKind.Heading1 => 28,
                ParagraphStyleKind.Heading2 => 22,
                ParagraphStyleKind.Heading3 => 18,
                _ => 14,
            },
            ListKind = listKind,
        };
    }

    private static TableBlock ConvertTable(Table table)
    {
        var result = new TableBlock();
        foreach (var row in table.Elements<TableRow>())
        {
            var rowBlock = new TableRowBlock();
            foreach (var cell in row.Elements<TableCell>())
            {
                var text = string.Join('\n',
                    cell.Elements<Paragraph>().Select(p =>
                        string.Concat(p.Descendants<Text>().Select(t => t.Text))));
                rowBlock.Cells.Add(new TableCellBlock { Text = text });
            }

            if (rowBlock.Cells.Count > 0)
                result.Rows.Add(rowBlock);
        }

        if (result.Rows.Count == 0)
            return TableBlock.Create(2, 2);

        return result;
    }

    private static void EnsureStyles(MainDocumentPart main)
    {
        var stylesPart = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();
        EnsureStyle(stylesPart.Styles, "Heading1", "heading 1", "32");
        EnsureStyle(stylesPart.Styles, "Heading2", "heading 2", "26");
        EnsureStyle(stylesPart.Styles, "Heading3", "heading 3", "22");
    }

    private static void EnsureStyle(Styles styles, string id, string name, string halfPoints)
    {
        if (styles.Elements<Style>().Any(s => s.StyleId == id))
            return;

        styles.AppendChild(new Style(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            new UIPriority { Val = 9 },
            new PrimaryStyle(),
            new StyleRunProperties(new Bold(), new FontSize { Val = halfPoints })
        )
        {
            Type = StyleValues.Paragraph,
            StyleId = id,
        });
    }

    private static void EnsureNumbering(MainDocumentPart main)
    {
        if (main.NumberingDefinitionsPart is not null)
            return;

        var part = main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = new Numbering(
            new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Bullet },
                    new LevelText { Val = "•" },
                    new LevelJustification { Val = LevelJustificationValues.Left }
                ) { LevelIndex = 0 }
            ) { AbstractNumberId = 1 },
            new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelJustification { Val = LevelJustificationValues.Left }
                ) { LevelIndex = 0 }
            ) { AbstractNumberId = 2 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
            new NumberingInstance(new AbstractNumId { Val = 2 }) { NumberID = 2 }
        );
    }

    private static Paragraph CreateParagraph(ParagraphBlock model, ref int bulletCount, ref int numberCount)
    {
        var pPr = new ParagraphProperties();

        pPr.AppendChild(model.Style switch
        {
            ParagraphStyleKind.Heading1 => new ParagraphStyleId { Val = "Heading1" },
            ParagraphStyleKind.Heading2 => new ParagraphStyleId { Val = "Heading2" },
            ParagraphStyleKind.Heading3 => new ParagraphStyleId { Val = "Heading3" },
            _ => new ParagraphStyleId { Val = "Normal" },
        });

        pPr.AppendChild(new Justification
        {
            Val = model.Alignment switch
            {
                ParagraphAlignmentKind.Center => JustificationValues.Center,
                ParagraphAlignmentKind.Right => JustificationValues.Right,
                ParagraphAlignmentKind.Justify => JustificationValues.Both,
                _ => JustificationValues.Left,
            },
        });

        if (model.ListKind == ListKind.Bullet)
        {
            pPr.AppendChild(new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = 1 }));
            bulletCount++;
        }
        else if (model.ListKind == ListKind.Numbered)
        {
            pPr.AppendChild(new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = 2 }));
            numberCount++;
        }

        var rPr = new RunProperties();
        if (model.IsBold)
            rPr.AppendChild(new Bold());
        if (model.IsItalic)
            rPr.AppendChild(new Italic());
        if (model.IsUnderline)
            rPr.AppendChild(new Underline { Val = UnderlineValues.Single });
        rPr.AppendChild(new FontSize { Val = ((int)Math.Round(model.FontSize * 2)).ToString() });
        rPr.AppendChild(new RunFonts
        {
            Ascii = "Microsoft JhengHei",
            EastAsia = "Microsoft JhengHei",
            HighAnsi = "Calibri",
        });

        var paragraph = new Paragraph(pPr);
        var lines = (model.Text ?? string.Empty).Replace("\r\n", "\n").Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
                paragraph.AppendChild(new Run(new Break()));
            paragraph.AppendChild(new Run((RunProperties)rPr.CloneNode(true), new Text(lines[i])
            {
                Space = SpaceProcessingModeValues.Preserve,
            }));
        }

        if (lines.Length == 0)
            paragraph.AppendChild(new Run((RunProperties)rPr.CloneNode(true), new Text(string.Empty)));

        return paragraph;
    }

    private static Table CreateTable(TableBlock model)
    {
        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "999999" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "999999" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "999999" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "999999" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "CCCCCC" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "CCCCCC" }
                )
            )
        );

        foreach (var row in model.Rows)
        {
            var tr = new TableRow();
            foreach (var cell in row.Cells)
            {
                var tc = new TableCell(
                    new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }),
                    new Paragraph(new Run(new Text(cell.Text ?? string.Empty)
                    {
                        Space = SpaceProcessingModeValues.Preserve,
                    }))
                );
                tr.AppendChild(tc);
            }

            table.AppendChild(tr);
        }

        return table;
    }
}
