using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocxAvalonia.Models;

namespace DocxAvalonia.Services;

/// <summary>Load / save editable document models as OpenDocument Text (.odt).</summary>
public sealed class OdtDocumentService
{
    private static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    private static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    private static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    private static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private static readonly XNamespace Draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    private static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
    private static readonly XNamespace Xlink = "http://www.w3.org/1999/xlink";
    private static readonly XNamespace Manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace Meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";

    public WordDocument Load(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("找不到指定的文件。", path);

        if (!string.Equals(Path.GetExtension(path), ".odt", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("僅支援 .odt 格式。");

        using var zip = ZipFile.OpenRead(path);
        var contentEntry = zip.GetEntry("content.xml")
            ?? throw new InvalidOperationException("無效的 ODT：缺少 content.xml。");

        XDocument contentDoc;
        using (var stream = contentEntry.Open())
            contentDoc = XDocument.Load(stream);

        var styles = LoadAutomaticStyles(contentDoc);
        var body = contentDoc.Root
            ?.Element(Office + "body")
            ?.Element(Office + "text")
            ?? throw new InvalidOperationException("無效的 ODT：缺少 office:body/office:text。");

        var doc = new WordDocument { Title = Path.GetFileNameWithoutExtension(path) };
        var pictures = LoadPictures(zip);

        foreach (var el in body.Elements())
        {
            if (el.Name == Text + "p" || el.Name == Text + "h")
            {
                foreach (var block in ConvertParagraphElement(el, styles, pictures))
                    doc.Blocks.Add(block);
            }
            else if (el.Name == Text + "list")
            {
                foreach (var item in el.Elements(Text + "list-item"))
                {
                    var kind = GuessListKind(el);
                    foreach (var p in item.Elements().Where(e => e.Name == Text + "p" || e.Name == Text + "h"))
                    {
                        foreach (var block in ConvertParagraphElement(p, styles, pictures, kind))
                            doc.Blocks.Add(block);
                    }
                }
            }
            else if (el.Name == Table + "table")
            {
                doc.Blocks.Add(ConvertTable(el, styles));
            }
            else if (el.Name == Text + "section")
            {
                foreach (var child in el.Elements())
                {
                    if (child.Name == Text + "p" || child.Name == Text + "h")
                    {
                        foreach (var block in ConvertParagraphElement(child, styles, pictures))
                            doc.Blocks.Add(block);
                    }
                    else if (child.Name == Table + "table")
                        doc.Blocks.Add(ConvertTable(child, styles));
                }
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

        var pictures = new List<(string Name, byte[] Bytes, string MediaType)>();
        var contentXml = BuildContentXml(model, pictures);
        var stylesXml = BuildStylesXml();
        var metaXml = BuildMetaXml(model.Title);
        var manifestXml = BuildManifestXml(pictures);

        using var fs = File.Create(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Create);

        // mimetype must be first and stored (no compression) per ODF.
        var mime = zip.CreateEntry("mimetype", CompressionLevel.NoCompression);
        using (var s = mime.Open())
        {
            var bytes = Encoding.UTF8.GetBytes("application/vnd.oasis.opendocument.text");
            s.Write(bytes, 0, bytes.Length);
        }

        WriteEntry(zip, "content.xml", contentXml);
        WriteEntry(zip, "styles.xml", stylesXml);
        WriteEntry(zip, "meta.xml", metaXml);
        WriteEntry(zip, "META-INF/manifest.xml", manifestXml);

        foreach (var pic in pictures)
        {
            var entry = zip.CreateEntry($"Pictures/{pic.Name}", CompressionLevel.Optimal);
            using var s = entry.Open();
            s.Write(pic.Bytes, 0, pic.Bytes.Length);
        }
    }

    // ─── Load helpers ───────────────────────────────────────

    private static Dictionary<string, StyleInfo> LoadAutomaticStyles(XDocument contentDoc)
    {
        var map = new Dictionary<string, StyleInfo>(StringComparer.Ordinal);
        var auto = contentDoc.Root?.Element(Office + "automatic-styles");
        if (auto is null)
            return map;

        foreach (var style in auto.Elements(Style + "style"))
        {
            var name = (string?)style.Attribute(Style + "name");
            if (string.IsNullOrEmpty(name))
                continue;

            var info = new StyleInfo();
            var parent = (string?)style.Attribute(Style + "parent-style-name");
            info.ParentStyle = parent;

            var tp = style.Element(Style + "text-properties");
            if (tp is not null)
            {
                info.Bold = IsBold(tp.Attribute(Fo + "font-weight")?.Value);
                info.Italic = IsItalic(tp.Attribute(Fo + "font-style")?.Value);
                info.Underline = IsUnderline(tp.Attribute(Style + "text-underline-style")?.Value);
                info.FontFamily = tp.Attribute(Style + "font-name")?.Value
                                  ?? tp.Attribute(Fo + "font-family")?.Value;
                info.FontSize = ParseFontSize(tp.Attribute(Fo + "font-size")?.Value);
            }

            var pp = style.Element(Style + "paragraph-properties");
            if (pp is not null)
                info.Alignment = MapAlign(pp.Attribute(Fo + "text-align")?.Value);

            map[name] = info;
        }

        return map;
    }

    private static Dictionary<string, byte[]> LoadPictures(ZipArchive zip)
    {
        var map = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.StartsWith("Pictures/", StringComparison.OrdinalIgnoreCase)
                && !entry.FullName.StartsWith("media/", StringComparison.OrdinalIgnoreCase))
                continue;
            if (entry.Length <= 0)
                continue;

            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            map[entry.FullName.Replace('\\', '/')] = ms.ToArray();
            map[Path.GetFileName(entry.FullName)] = ms.ToArray();
        }

        return map;
    }

    private static IEnumerable<DocumentBlock> ConvertParagraphElement(
        XElement el,
        Dictionary<string, StyleInfo> styles,
        Dictionary<string, byte[]> pictures,
        ListKind listKind = ListKind.None)
    {
        var images = new List<ImageBlock>();
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

        void Walk(XElement node, StyleInfo? inherited = null)
        {
            foreach (var child in node.Nodes())
            {
                if (child is XText xt)
                {
                    var t = xt.Value;
                    if (string.IsNullOrEmpty(t))
                        continue;
                    sb.Append(t);
                    runCount++;
                    if (inherited?.Bold == true) boldCount++;
                    if (inherited?.Italic == true) italicCount++;
                    if (inherited?.Underline == true) underlineCount++;
                    fontSize ??= inherited?.FontSize;
                    fontFamily ??= inherited?.FontFamily;
                }
                else if (child is XElement xe)
                {
                    if (xe.Name == Text + "span")
                    {
                        StyleInfo? spanStyle = inherited;
                        var sn = (string?)xe.Attribute(Text + "style-name");
                        if (!string.IsNullOrEmpty(sn) && styles.TryGetValue(sn, out var st))
                            spanStyle = st;
                        Walk(xe, spanStyle);
                    }
                    else if (xe.Name == Text + "s")
                    {
                        var c = (int?)xe.Attribute(Text + "c") ?? 1;
                        sb.Append(new string(' ', Math.Clamp(c, 1, 200)));
                    }
                    else if (xe.Name == Text + "tab")
                        sb.Append('\t');
                    else if (xe.Name == Text + "line-break")
                        sb.AppendLine();
                    else if (xe.Name == Text + "a")
                        Walk(xe, inherited);
                    else if (xe.Name == Draw + "frame")
                    {
                        var img = TryExtractImage(xe, pictures);
                        if (img is not null)
                            images.Add(img);
                    }
                    else
                        Walk(xe, inherited);
                }
            }
        }

        var paraStyleName = (string?)el.Attribute(Text + "style-name");
        StyleInfo? paraStyle = null;
        if (!string.IsNullOrEmpty(paraStyleName) && styles.TryGetValue(paraStyleName, out var ps))
            paraStyle = ps;

        Walk(el, paraStyle);

        if (runCount > 0)
        {
            bold = boldCount * 2 >= runCount;
            italic = italicCount * 2 >= runCount;
            underline = underlineCount * 2 >= runCount;
        }

        var styleKind = ParagraphStyleKind.Normal;
        if (el.Name == Text + "h")
        {
            var level = (int?)el.Attribute(Text + "outline-level") ?? 1;
            styleKind = level switch
            {
                1 => ParagraphStyleKind.Heading1,
                2 => ParagraphStyleKind.Heading2,
                _ => ParagraphStyleKind.Heading3,
            };
        }
        else if (!string.IsNullOrEmpty(paraStyle?.ParentStyle))
        {
            styleKind = MapHeadingName(paraStyle.ParentStyle);
        }
        else if (!string.IsNullOrEmpty(paraStyleName))
        {
            styleKind = MapHeadingName(paraStyleName);
        }

        var alignment = paraStyle?.Alignment ?? ParagraphAlignmentKind.Left;

        if (!string.IsNullOrEmpty(sb.ToString()) || images.Count == 0)
        {
            yield return new ParagraphBlock
            {
                Text = sb.ToString(),
                Style = styleKind,
                Alignment = alignment,
                IsBold = bold || styleKind != ParagraphStyleKind.Normal,
                IsItalic = italic,
                IsUnderline = underline,
                FontSize = fontSize ?? styleKind switch
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

        foreach (var img in images)
            yield return img;
    }

    private static ImageBlock? TryExtractImage(XElement frame, Dictionary<string, byte[]> pictures)
    {
        var image = frame.Descendants(Draw + "image").FirstOrDefault();
        var href = (string?)image?.Attribute(Xlink + "href");
        if (string.IsNullOrEmpty(href))
            return null;

        var key = href.Replace('\\', '/').TrimStart('.');
        if (key.StartsWith('/'))
            key = key.TrimStart('/');
        if (!pictures.TryGetValue(key, out var bytes)
            && !pictures.TryGetValue(Path.GetFileName(key), out bytes))
            return null;

        if (bytes.Length == 0)
            return null;

        double widthDip = 400;
        double heightDip = 0;
        var w = (string?)frame.Attribute(Svg + "width");
        var h = (string?)frame.Attribute(Svg + "height");
        if (TryParseLengthToDip(w, out var wd))
            widthDip = Math.Clamp(wd, 40, 720);
        if (TryParseLengthToDip(h, out var hd))
            heightDip = Math.Max(0, hd);

        var ext = Path.GetExtension(key).ToLowerInvariant();
        var contentType = ext switch
        {
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".webp" => "image/webp",
            _ => "image/png",
        };

        return new ImageBlock
        {
            ImageBytes = bytes,
            ContentType = contentType,
            FileName = Path.GetFileName(key),
            DisplayWidth = widthDip,
            DisplayHeight = heightDip,
        };
    }

    private static TableBlock ConvertTable(XElement table, Dictionary<string, StyleInfo> styles)
    {
        var result = new TableBlock();
        foreach (var row in table.Elements(Table + "table-row"))
        {
            var rowBlock = new TableRowBlock();
            foreach (var cell in row.Elements(Table + "table-cell"))
            {
                var textParts = new List<string>();
                double fontSize = 14;
                var bold = false;
                var italic = false;
                var underline = false;
                string? fontFamily = null;
                var alignment = ParagraphAlignmentKind.Left;
                var first = true;

                foreach (var p in cell.Elements().Where(e => e.Name == Text + "p" || e.Name == Text + "h"))
                {
                    var blocks = ConvertParagraphElement(p, styles, new Dictionary<string, byte[]>()).OfType<ParagraphBlock>().ToList();
                    foreach (var b in blocks)
                    {
                        textParts.Add(b.Text);
                        if (first)
                        {
                            fontSize = b.FontSize;
                            bold = b.IsBold;
                            italic = b.IsItalic;
                            underline = b.IsUnderline;
                            fontFamily = b.FontFamily;
                            alignment = b.Alignment;
                            first = false;
                        }
                    }
                }

                // Covered cells / empty
                var cols = (int?)cell.Attribute(Table + "number-columns-spanned") ?? 1;
                rowBlock.Cells.Add(new TableCellBlock
                {
                    Text = string.Join('\n', textParts),
                    FontSize = fontSize,
                    IsBold = bold,
                    IsItalic = italic,
                    IsUnderline = underline,
                    FontFamily = fontFamily ?? "Microsoft JhengHei",
                    Alignment = alignment,
                });
                for (var i = 1; i < cols; i++)
                    rowBlock.Cells.Add(new TableCellBlock());
            }

            if (rowBlock.Cells.Count > 0)
                result.Rows.Add(rowBlock);
        }

        if (result.Rows.Count == 0)
            return TableBlock.Create(2, 2);
        return result;
    }

    private static ListKind GuessListKind(XElement list)
    {
        var styleName = (string?)list.Attribute(Text + "style-name") ?? "";
        if (styleName.Contains("Number", StringComparison.OrdinalIgnoreCase)
            || styleName.Contains("編號", StringComparison.OrdinalIgnoreCase))
            return ListKind.Numbered;
        return ListKind.Bullet;
    }

    // ─── Save helpers ───────────────────────────────────────

    private static string BuildContentXml(WordDocument model, List<(string Name, byte[] Bytes, string MediaType)> pictures)
    {
        var autoStyles = new XElement(Office + "automatic-styles");
        var textBody = new XElement(Office + "text");
        var styleIndex = 0;
        var imageIndex = 0;

        foreach (var block in model.Blocks)
        {
            switch (block)
            {
                case ParagraphBlock p:
                    var styleName = EnsureParaStyle(autoStyles, p, ref styleIndex);
                    var pEl = p.Style switch
                    {
                        ParagraphStyleKind.Heading1 => new XElement(Text + "h",
                            new XAttribute(Text + "outline-level", 1),
                            new XAttribute(Text + "style-name", styleName)),
                        ParagraphStyleKind.Heading2 => new XElement(Text + "h",
                            new XAttribute(Text + "outline-level", 2),
                            new XAttribute(Text + "style-name", styleName)),
                        ParagraphStyleKind.Heading3 => new XElement(Text + "h",
                            new XAttribute(Text + "outline-level", 3),
                            new XAttribute(Text + "style-name", styleName)),
                        _ => new XElement(Text + "p", new XAttribute(Text + "style-name", styleName)),
                    };

                    AppendTextWithBreaks(pEl, p.Text);

                    if (p.ListKind is ListKind.Bullet or ListKind.Numbered)
                    {
                        var list = new XElement(Text + "list",
                            new XAttribute(Text + "style-name", p.ListKind == ListKind.Numbered ? "LNumber" : "LBullet"),
                            new XElement(Text + "list-item", pEl));
                        textBody.Add(list);
                    }
                    else
                    {
                        textBody.Add(pEl);
                    }

                    break;

                case TableBlock t:
                    textBody.Add(BuildTableElement(t, autoStyles, ref styleIndex));
                    break;

                case ImageBlock img:
                    if (img.ImageBytes is not { Length: > 0 })
                        break;
                    imageIndex++;
                    var ext = GuessImageExtension(img.ContentType, img.FileName);
                    var fileName = $"image{imageIndex}{ext}";
                    var media = img.ContentType;
                    if (string.IsNullOrWhiteSpace(media))
                        media = "image/png";
                    pictures.Add((fileName, img.ImageBytes, media));

                    var widthCm = Math.Clamp(img.DisplayWidth, 40, 720) / 96.0 * 2.54;
                    var heightCm = (img.DisplayHeight > 0 ? img.DisplayHeight : img.ResolvedHeight) / 96.0 * 2.54;
                    heightCm = Math.Clamp(heightCm, 0.5, 40);
                    widthCm = Math.Clamp(widthCm, 0.5, 20);

                    textBody.Add(new XElement(Text + "p",
                        new XElement(Draw + "frame",
                            new XAttribute(Draw + "name", $"Image{imageIndex}"),
                            new XAttribute(Svg + "width", FormattableString.Invariant($"{widthCm:0.###}cm")),
                            new XAttribute(Svg + "height", FormattableString.Invariant($"{heightCm:0.###}cm")),
                            new XAttribute(Text + "anchor-type", "paragraph"),
                            new XElement(Draw + "image",
                                new XAttribute(Xlink + "href", $"Pictures/{fileName}"),
                                new XAttribute(Xlink + "type", "simple"),
                                new XAttribute(Xlink + "show", "embed"),
                                new XAttribute(Xlink + "actuate", "onLoad")))));
                    break;
            }
        }

        if (!textBody.HasElements)
            textBody.Add(new XElement(Text + "p"));

        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(Office + "document-content",
                new XAttribute(XNamespace.Xmlns + "office", Office),
                new XAttribute(XNamespace.Xmlns + "text", Text),
                new XAttribute(XNamespace.Xmlns + "style", Style),
                new XAttribute(XNamespace.Xmlns + "fo", Fo),
                new XAttribute(XNamespace.Xmlns + "table", Table),
                new XAttribute(XNamespace.Xmlns + "draw", Draw),
                new XAttribute(XNamespace.Xmlns + "svg", Svg),
                new XAttribute(XNamespace.Xmlns + "xlink", Xlink),
                new XAttribute(Office + "version", "1.2"),
                autoStyles,
                new XElement(Office + "body", textBody)));

        return doc.ToString(SaveOptions.DisableFormatting);
    }

    private static XElement BuildTableElement(TableBlock table, XElement autoStyles, ref int styleIndex)
    {
        table.EnsureRectangular();
        var tEl = new XElement(Table + "table", new XAttribute(Table + "name", "Table1"));
        var cols = Math.Max(1, table.ColumnCount);
        for (var c = 0; c < cols; c++)
            tEl.Add(new XElement(Table + "table-column"));

        foreach (var row in table.Rows)
        {
            var rEl = new XElement(Table + "table-row");
            foreach (var cell in row.Cells)
            {
                var styleName = EnsureCellStyle(autoStyles, cell, ref styleIndex);
                var cEl = new XElement(Table + "table-cell",
                    new XAttribute(Office + "value-type", "string"),
                    new XElement(Text + "p",
                        new XAttribute(Text + "style-name", styleName),
                        cell.Text ?? string.Empty));
                rEl.Add(cEl);
            }

            tEl.Add(rEl);
        }

        return tEl;
    }

    private static string EnsureParaStyle(XElement autoStyles, ParagraphBlock p, ref int styleIndex)
    {
        styleIndex++;
        var name = $"P{styleIndex}";
        var parent = p.Style switch
        {
            ParagraphStyleKind.Heading1 => "Heading_20_1",
            ParagraphStyleKind.Heading2 => "Heading_20_2",
            ParagraphStyleKind.Heading3 => "Heading_20_3",
            _ => "Standard",
        };

        autoStyles.Add(new XElement(Style + "style",
            new XAttribute(Style + "name", name),
            new XAttribute(Style + "family", "paragraph"),
            new XAttribute(Style + "parent-style-name", parent),
            new XElement(Style + "paragraph-properties",
                new XAttribute(Fo + "text-align", MapAlignOut(p.Alignment))),
            new XElement(Style + "text-properties",
                new XAttribute(Style + "font-name", p.FontFamily ?? "Microsoft JhengHei"),
                new XAttribute(Fo + "font-size", FormattableString.Invariant($"{p.FontSize}pt")),
                new XAttribute(Fo + "font-weight", p.IsBold ? "bold" : "normal"),
                new XAttribute(Fo + "font-style", p.IsItalic ? "italic" : "normal"),
                p.IsUnderline
                    ? new XAttribute(Style + "text-underline-style", "solid")
                    : null,
                p.IsUnderline
                    ? new XAttribute(Style + "text-underline-width", "auto")
                    : null)));

        return name;
    }

    private static string EnsureCellStyle(XElement autoStyles, TableCellBlock cell, ref int styleIndex)
    {
        styleIndex++;
        var name = $"C{styleIndex}";
        autoStyles.Add(new XElement(Style + "style",
            new XAttribute(Style + "name", name),
            new XAttribute(Style + "family", "paragraph"),
            new XElement(Style + "paragraph-properties",
                new XAttribute(Fo + "text-align", MapAlignOut(cell.Alignment))),
            new XElement(Style + "text-properties",
                new XAttribute(Style + "font-name", cell.FontFamily ?? "Microsoft JhengHei"),
                new XAttribute(Fo + "font-size", FormattableString.Invariant($"{cell.FontSize}pt")),
                new XAttribute(Fo + "font-weight", cell.IsBold ? "bold" : "normal"),
                new XAttribute(Fo + "font-style", cell.IsItalic ? "italic" : "normal"),
                cell.IsUnderline
                    ? new XAttribute(Style + "text-underline-style", "solid")
                    : null)));
        return name;
    }

    private static string BuildStylesXml()
    {
        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(Office + "document-styles",
                new XAttribute(XNamespace.Xmlns + "office", Office),
                new XAttribute(XNamespace.Xmlns + "style", Style),
                new XAttribute(XNamespace.Xmlns + "text", Text),
                new XAttribute(XNamespace.Xmlns + "fo", Fo),
                new XAttribute(Office + "version", "1.2"),
                new XElement(Office + "styles",
                    new XElement(Style + "style",
                        new XAttribute(Style + "name", "Standard"),
                        new XAttribute(Style + "family", "paragraph"),
                        new XAttribute(Style + "class", "text")),
                    new XElement(Style + "style",
                        new XAttribute(Style + "name", "Heading_20_1"),
                        new XAttribute(Style + "display-name", "Heading 1"),
                        new XAttribute(Style + "family", "paragraph"),
                        new XAttribute(Style + "parent-style-name", "Standard"),
                        new XAttribute(Style + "class", "text"),
                        new XElement(Style + "text-properties",
                            new XAttribute(Fo + "font-size", "28pt"),
                            new XAttribute(Fo + "font-weight", "bold"))),
                    new XElement(Style + "style",
                        new XAttribute(Style + "name", "Heading_20_2"),
                        new XAttribute(Style + "display-name", "Heading 2"),
                        new XAttribute(Style + "family", "paragraph"),
                        new XAttribute(Style + "parent-style-name", "Standard"),
                        new XAttribute(Style + "class", "text"),
                        new XElement(Style + "text-properties",
                            new XAttribute(Fo + "font-size", "22pt"),
                            new XAttribute(Fo + "font-weight", "bold"))),
                    new XElement(Style + "style",
                        new XAttribute(Style + "name", "Heading_20_3"),
                        new XAttribute(Style + "display-name", "Heading 3"),
                        new XAttribute(Style + "family", "paragraph"),
                        new XAttribute(Style + "parent-style-name", "Standard"),
                        new XAttribute(Style + "class", "text"),
                        new XElement(Style + "text-properties",
                            new XAttribute(Fo + "font-size", "18pt"),
                            new XAttribute(Fo + "font-weight", "bold"))),
                    new XElement(Text + "list-style",
                        new XAttribute(Style + "name", "LBullet"),
                        new XElement(Text + "list-level-style-bullet",
                            new XAttribute(Text + "level", "1"),
                            new XAttribute(Text + "bullet-char", "•"))),
                    new XElement(Text + "list-style",
                        new XAttribute(Style + "name", "LNumber"),
                        new XElement(Text + "list-level-style-number",
                            new XAttribute(Text + "level", "1"),
                            new XAttribute(Style + "num-format", "1"),
                            new XAttribute(Style + "num-suffix", "."))))));
        return doc.ToString(SaveOptions.DisableFormatting);
    }

    private static string BuildMetaXml(string? title)
    {
        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(Office + "document-meta",
                new XAttribute(XNamespace.Xmlns + "office", Office),
                new XAttribute(XNamespace.Xmlns + "dc", Dc),
                new XAttribute(XNamespace.Xmlns + "meta", Meta),
                new XAttribute(Office + "version", "1.2"),
                new XElement(Office + "meta",
                    new XElement(Dc + "title", title ?? "文件"),
                    new XElement(Meta + "generator", "DocxAvalonia"))));
        return doc.ToString(SaveOptions.DisableFormatting);
    }

    private static string BuildManifestXml(List<(string Name, byte[] Bytes, string MediaType)> pictures)
    {
        var root = new XElement(Manifest + "manifest",
            new XAttribute(XNamespace.Xmlns + "manifest", Manifest),
            new XAttribute(Manifest + "version", "1.2"),
            new XElement(Manifest + "file-entry",
                new XAttribute(Manifest + "full-path", "/"),
                new XAttribute(Manifest + "version", "1.2"),
                new XAttribute(Manifest + "media-type", "application/vnd.oasis.opendocument.text")),
            new XElement(Manifest + "file-entry",
                new XAttribute(Manifest + "full-path", "content.xml"),
                new XAttribute(Manifest + "media-type", "text/xml")),
            new XElement(Manifest + "file-entry",
                new XAttribute(Manifest + "full-path", "styles.xml"),
                new XAttribute(Manifest + "media-type", "text/xml")),
            new XElement(Manifest + "file-entry",
                new XAttribute(Manifest + "full-path", "meta.xml"),
                new XAttribute(Manifest + "media-type", "text/xml")));

        foreach (var pic in pictures)
        {
            root.Add(new XElement(Manifest + "file-entry",
                new XAttribute(Manifest + "full-path", $"Pictures/{pic.Name}"),
                new XAttribute(Manifest + "media-type", pic.MediaType)));
        }

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
        return doc.ToString(SaveOptions.DisableFormatting);
    }

    private static void WriteEntry(ZipArchive zip, string name, string xml)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var s = entry.Open();
        using var w = new StreamWriter(s, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        w.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        w.Write(xml);
    }

    private static void AppendTextWithBreaks(XElement parent, string? text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        var parts = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (var i = 0; i < parts.Length; i++)
        {
            if (i > 0)
                parent.Add(new XElement(Text + "line-break"));
            if (parts[i].Length > 0)
                parent.Add(new XText(parts[i]));
        }
    }

    // ─── Shared utils ───────────────────────────────────────

    private static ParagraphStyleKind MapHeadingName(string? name)
    {
        if (string.IsNullOrEmpty(name))
            return ParagraphStyleKind.Normal;
        var n = name.Replace("_20_", " ", StringComparison.Ordinal).ToLowerInvariant();
        if (n.Contains("heading 1") || n is "heading1" or "title" or "標題1" or "標題 1")
            return ParagraphStyleKind.Heading1;
        if (n.Contains("heading 2") || n is "heading2" or "subtitle" or "標題2" or "標題 2")
            return ParagraphStyleKind.Heading2;
        if (n.Contains("heading 3") || n is "heading3" or "標題3" or "標題 3")
            return ParagraphStyleKind.Heading3;
        return ParagraphStyleKind.Normal;
    }

    private static ParagraphAlignmentKind MapAlign(string? align) =>
        align?.ToLowerInvariant() switch
        {
            "center" => ParagraphAlignmentKind.Center,
            "end" or "right" => ParagraphAlignmentKind.Right,
            "justify" => ParagraphAlignmentKind.Justify,
            _ => ParagraphAlignmentKind.Left,
        };

    private static string MapAlignOut(ParagraphAlignmentKind a) =>
        a switch
        {
            ParagraphAlignmentKind.Center => "center",
            ParagraphAlignmentKind.Right => "end",
            ParagraphAlignmentKind.Justify => "justify",
            _ => "start",
        };

    private static bool IsBold(string? v) =>
        v is not null && (v.Equals("bold", StringComparison.OrdinalIgnoreCase) || v == "700" || v == "800" || v == "900");

    private static bool IsItalic(string? v) =>
        v is not null && v.Equals("italic", StringComparison.OrdinalIgnoreCase);

    private static bool IsUnderline(string? v) =>
        v is not null && !v.Equals("none", StringComparison.OrdinalIgnoreCase) && v.Length > 0;

    private static double? ParseFontSize(string? v)
    {
        if (string.IsNullOrWhiteSpace(v))
            return null;
        v = v.Trim();
        if (v.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
            && double.TryParse(v[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var pt))
            return pt;
        if (v.EndsWith("pc", StringComparison.OrdinalIgnoreCase)
            && double.TryParse(v[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var pc))
            return pc * 12;
        if (double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
            return n;
        return null;
    }

    private static bool TryParseLengthToDip(string? v, out double dip)
    {
        dip = 0;
        if (string.IsNullOrWhiteSpace(v))
            return false;
        v = v.Trim();
        double value;
        if (v.EndsWith("cm", StringComparison.OrdinalIgnoreCase)
            && double.TryParse(v[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out value))
        {
            dip = value / 2.54 * 96.0;
            return true;
        }

        if (v.EndsWith("in", StringComparison.OrdinalIgnoreCase)
            && double.TryParse(v[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out value))
        {
            dip = value * 96.0;
            return true;
        }

        if (v.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
            && double.TryParse(v[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out value))
        {
            dip = value * 96.0 / 72.0;
            return true;
        }

        if (v.EndsWith("px", StringComparison.OrdinalIgnoreCase)
            && double.TryParse(v[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out value))
        {
            dip = value;
            return true;
        }

        return false;
    }

    private static string GuessImageExtension(string? contentType, string? fileName)
    {
        var fromName = Path.GetExtension(fileName ?? "");
        if (!string.IsNullOrEmpty(fromName))
            return fromName.ToLowerInvariant();

        return (contentType ?? "").ToLowerInvariant() switch
        {
            "image/jpeg" or "image/jpg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/webp" => ".webp",
            _ => ".png",
        };
    }

    private sealed class StyleInfo
    {
        public string? ParentStyle;
        public bool? Bold;
        public bool? Italic;
        public bool? Underline;
        public double? FontSize;
        public string? FontFamily;
        public ParagraphAlignmentKind Alignment = ParagraphAlignmentKind.Left;
    }
}
