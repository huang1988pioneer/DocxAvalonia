using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxAvalonia.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxAvalonia.Services;

/// <summary>Load / save editable document models as .docx (text, tables, images).</summary>
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
                    foreach (var block in ConvertParagraphWithImages(paragraph, numbering, main))
                        doc.Blocks.Add(block);
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
        var imageIndex = 0;

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
                case ImageBlock img:
                    body.AppendChild(CreateImageParagraph(main, img, ref imageIndex));
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

    private static IEnumerable<DocumentBlock> ConvertParagraphWithImages(
        Paragraph paragraph,
        Dictionary<int, ListKind> numbering,
        MainDocumentPart main)
    {
        var images = ExtractImages(paragraph, main).ToList();
        var textBlock = ConvertParagraphText(paragraph, numbering);

        // Emit text when there is content, or when there are no images (preserve blank lines).
        if (!string.IsNullOrEmpty(textBlock.Text) || images.Count == 0)
            yield return textBlock;

        foreach (var image in images)
            yield return image;
    }

    private static ParagraphBlock ConvertParagraphText(Paragraph paragraph, Dictionary<int, ListKind> numbering)
    {
        var sb = new StringBuilder();
        var bold = false;
        var italic = false;
        var underline = false;
        double? fontSize = null;
        string? fontFamily = null;
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

            fontFamily ??= ResolveFontFamilyName(props?.RunFonts);
        }

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
            FontFamily = fontFamily ?? "Microsoft JhengHei",
            ListKind = listKind,
        };
    }

    private static string? ResolveFontFamilyName(RunFonts? fonts)
    {
        if (fonts is null)
            return null;
        return fonts.EastAsia?.Value
               ?? fonts.Ascii?.Value
               ?? fonts.HighAnsi?.Value
               ?? fonts.ComplexScript?.Value;
    }

    private static IEnumerable<ImageBlock> ExtractImages(Paragraph paragraph, MainDocumentPart main)
    {
        var seenRelIds = new HashSet<string>(StringComparer.Ordinal);

        // DrawingML images (modern Word): inline + floating (anchor)
        foreach (var drawing in paragraph.Descendants<Drawing>())
        {
            ImageBlock? block = null;
            try
            {
                block = TryExtractDrawingImage(drawing, main, seenRelIds);
            }
            catch
            {
                // Never let a single broken drawing crash the whole document load.
            }

            if (block is not null)
                yield return block;
        }

        // VML fallback (older Word) — only if no DrawingML images found in this paragraph
        if (seenRelIds.Count > 0)
            yield break;

        IEnumerable<DocumentFormat.OpenXml.Vml.ImageData> vmlList;
        try
        {
            vmlList = paragraph.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().ToList();
        }
        catch
        {
            yield break;
        }

        foreach (var vml in vmlList)
        {
            ImageBlock? block = null;
            try
            {
                var relId = vml.RelationshipId?.Value;
                if (string.IsNullOrEmpty(relId) || !seenRelIds.Add(relId))
                    continue;
                block = TryLoadImagePart(main, relId, 400, 0);
            }
            catch
            {
                // ignore
            }

            if (block is not null)
                yield return block;
        }
    }

    private static ImageBlock? TryExtractDrawingImage(Drawing drawing, MainDocumentPart main, HashSet<string> seenRelIds)
    {
        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        // Prefer embedded; skip external links (often not available offline).
        var embedId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId))
            return null;
        if (!seenRelIds.Add(embedId))
            return null;

        double widthDip = 400;
        double heightDip = 0;
        try
        {
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            if (extent?.Cx?.Value is { } cx && cx > 0)
            {
                widthDip = cx / 914400.0 * 96.0;
                if (extent.Cy?.Value is { } cy && cy > 0)
                    heightDip = cy / 914400.0 * 96.0;
            }
        }
        catch
        {
            // keep defaults
        }

        if (double.IsNaN(widthDip) || double.IsInfinity(widthDip) || widthDip <= 0)
            widthDip = 400;
        widthDip = Math.Clamp(widthDip, 40, 720);
        if (double.IsNaN(heightDip) || double.IsInfinity(heightDip) || heightDip < 0)
            heightDip = 0;

        return TryLoadImagePart(main, embedId, widthDip, heightDip);
    }

    private static ImageBlock? TryLoadImagePart(MainDocumentPart main, string relId, double widthDip, double heightDip)
    {
        OpenXmlPart? part;
        try
        {
            part = main.GetPartById(relId);
        }
        catch
        {
            return null;
        }

        if (part is not ImagePart imagePart)
            return null;

        byte[] bytes;
        try
        {
            using var s = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            bytes = ms.ToArray();
        }
        catch
        {
            return null;
        }

        if (bytes.Length == 0)
            return null;

        string? fileName = null;
        try
        {
            fileName = Path.GetFileName(imagePart.Uri?.OriginalString ?? "image");
        }
        catch
        {
            fileName = "image";
        }

        return new ImageBlock
        {
            ImageBytes = bytes,
            ContentType = imagePart.ContentType ?? "image/png",
            FileName = fileName,
            DisplayWidth = widthDip,
            DisplayHeight = heightDip,
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
                rowBlock.Cells.Add(ConvertTableCell(cell));
            }

            if (rowBlock.Cells.Count > 0)
                result.Rows.Add(rowBlock);
        }

        if (result.Rows.Count == 0)
            return TableBlock.Create(2, 2);

        return result;
    }

    private static TableCellBlock ConvertTableCell(TableCell cell)
    {
        var paragraphs = cell.Elements<Paragraph>().ToList();
        var text = string.Join('\n',
            paragraphs.Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))));

        // Sample formatting from first non-empty run in the cell.
        double fontSize = 14;
        var bold = false;
        var italic = false;
        var underline = false;
        string? fontFamily = null;
        var alignment = ParagraphAlignmentKind.Left;

        var firstPara = paragraphs.FirstOrDefault();
        if (firstPara is not null)
        {
            var align = firstPara.ParagraphProperties?.Justification?.Val?.Value;
            if (align == JustificationValues.Center)
                alignment = ParagraphAlignmentKind.Center;
            else if (align == JustificationValues.Right)
                alignment = ParagraphAlignmentKind.Right;
            else if (align == JustificationValues.Both)
                alignment = ParagraphAlignmentKind.Justify;

            var runs = firstPara.Descendants<Run>().Where(r => r.Elements<Text>().Any()).ToList();
            if (runs.Count > 0)
            {
                var props = runs[0].RunProperties;
                if (props?.Bold is not null && (props.Bold.Val is null || props.Bold.Val.Value))
                    bold = true;
                if (props?.Italic is not null && (props.Italic.Val is null || props.Italic.Val.Value))
                    italic = true;
                if (props?.Underline is not null
                    && props.Underline.Val is not null
                    && props.Underline.Val.Value != UnderlineValues.None)
                    underline = true;
                var half = props?.FontSize?.Val?.Value;
                if (half is not null && double.TryParse(half, out var hp) && hp > 0)
                    fontSize = hp / 2.0;
                fontFamily = ResolveFontFamilyName(props?.RunFonts);
            }
        }

        return new TableCellBlock
        {
            Text = text,
            FontSize = fontSize,
            FontFamily = fontFamily ?? "Microsoft JhengHei",
            IsBold = bold,
            IsItalic = italic,
            IsUnderline = underline,
            Alignment = alignment,
        };
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
        var fontName = string.IsNullOrWhiteSpace(model.FontFamily) ? "Microsoft JhengHei" : model.FontFamily.Trim();
        rPr.AppendChild(CreateRunFonts(fontName));

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

    private static RunFonts CreateRunFonts(string fontName)
    {
        // Latin-friendly fonts use themselves for ASCII/HighAnsi; CJK names map East Asia primarily.
        var isCjk = fontName.Contains("JhengHei", StringComparison.OrdinalIgnoreCase)
                    || fontName.Contains("YaHei", StringComparison.OrdinalIgnoreCase)
                    || fontName.Contains("Ming", StringComparison.OrdinalIgnoreCase)
                    || fontName.Contains("Song", StringComparison.OrdinalIgnoreCase)
                    || fontName.Contains("Kai", StringComparison.OrdinalIgnoreCase)
                    || fontName.Contains("Gothic", StringComparison.OrdinalIgnoreCase)
                    || fontName.Contains("Noto Sans CJK", StringComparison.OrdinalIgnoreCase)
                    || fontName.Contains("Noto Sans TC", StringComparison.OrdinalIgnoreCase)
                    || fontName.Contains("Source Han", StringComparison.OrdinalIgnoreCase)
                    || fontName.Contains("蘋方", StringComparison.Ordinal)
                    || fontName.Contains("微软", StringComparison.Ordinal)
                    || fontName.Contains("微軟", StringComparison.Ordinal);

        if (isCjk)
        {
            return new RunFonts
            {
                Ascii = fontName,
                EastAsia = fontName,
                HighAnsi = "Calibri",
            };
        }

        return new RunFonts
        {
            Ascii = fontName,
            EastAsia = "Microsoft JhengHei",
            HighAnsi = fontName,
        };
    }

    private static Paragraph CreateImageParagraph(MainDocumentPart main, ImageBlock model, ref int imageIndex)
    {
        imageIndex++;
        var partType = model.ContentType switch
        {
            "image/jpeg" or "image/jpg" => ImagePartType.Jpeg,
            "image/gif" => ImagePartType.Gif,
            "image/bmp" => ImagePartType.Bmp,
            "image/tiff" => ImagePartType.Tiff,
            "image/x-emf" => ImagePartType.Emf,
            "image/x-wmf" => ImagePartType.Wmf,
            _ => ImagePartType.Png,
        };

        var imagePart = main.AddImagePart(partType);
        using (var ms = new MemoryStream(model.ImageBytes))
            imagePart.FeedData(ms);

        var relId = main.GetIdOfPart(imagePart);

        // DIP (96) → EMU — never touch Avalonia Bitmap here (may run off UI thread).
        var widthDip = model.DisplayWidth > 0 && !double.IsNaN(model.DisplayWidth)
            ? model.DisplayWidth
            : 400;
        var heightDip = model.DisplayHeight;
        if ((heightDip <= 0 || double.IsNaN(heightDip)) && model.PixelWidth > 0 && model.PixelHeight > 0)
            heightDip = widthDip * model.PixelHeight / model.PixelWidth;
        if (heightDip <= 0 || double.IsNaN(heightDip))
            heightDip = widthDip * 0.75;
        widthDip = Math.Clamp(widthDip, 40, 720);
        heightDip = Math.Clamp(heightDip, 20, 2000);

        var cx = (long)(widthDip / 96.0 * 914400);
        var cy = (long)(heightDip / 96.0 * 914400);

        var element =
            new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = cx, Cy = cy },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = (uint)imageIndex, Name = model.FileName ?? $"Picture {imageIndex}" },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties
                                    {
                                        Id = (uint)imageIndex,
                                        Name = model.FileName ?? $"image{imageIndex}",
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = cx, Cy = cy }),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    {
                                        Preset = A.ShapeTypeValues.Rectangle,
                                    }))
                        )
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                        })
                )
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U,
                });

        return new Paragraph(new Run(element));
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
                tr.AppendChild(CreateTableCell(cell));
            }

            table.AppendChild(tr);
        }

        return table;
    }

    private static TableCell CreateTableCell(TableCellBlock cell)
    {
        var pPr = new ParagraphProperties(
            new Justification
            {
                Val = cell.Alignment switch
                {
                    ParagraphAlignmentKind.Center => JustificationValues.Center,
                    ParagraphAlignmentKind.Right => JustificationValues.Right,
                    ParagraphAlignmentKind.Justify => JustificationValues.Both,
                    _ => JustificationValues.Left,
                },
            });

        var rPr = new RunProperties();
        if (cell.IsBold)
            rPr.AppendChild(new Bold());
        if (cell.IsItalic)
            rPr.AppendChild(new Italic());
        if (cell.IsUnderline)
            rPr.AppendChild(new Underline { Val = UnderlineValues.Single });
        var size = cell.FontSize > 0 ? cell.FontSize : 14;
        rPr.AppendChild(new FontSize { Val = ((int)Math.Round(size * 2)).ToString() });
        var fontName = string.IsNullOrWhiteSpace(cell.FontFamily) ? "Microsoft JhengHei" : cell.FontFamily.Trim();
        rPr.AppendChild(CreateRunFonts(fontName));

        var paragraph = new Paragraph(pPr);
        var lines = (cell.Text ?? string.Empty).Replace("\r\n", "\n").Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
                paragraph.AppendChild(new Run(new Break()));
            paragraph.AppendChild(new Run(
                (RunProperties)rPr.CloneNode(true),
                new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve }));
        }

        if (lines.Length == 0)
            paragraph.AppendChild(new Run((RunProperties)rPr.CloneNode(true), new Text(string.Empty)));

        return new TableCell(
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }),
            paragraph);
    }
}
