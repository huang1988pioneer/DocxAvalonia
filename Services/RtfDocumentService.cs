using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using DocxAvalonia.Models;

namespace DocxAvalonia.Services;

/// <summary>
/// Minimal RTF read/write for paragraphs, basic formatting, tables, and PNG/JPEG images.
/// Used as the on-disk representation for .doc save/load (Word/WPS/LibreOffice compatible).
/// </summary>
public sealed class RtfDocumentService
{
    public WordDocument Load(string path)
    {
        var rtf = File.ReadAllText(path, Encoding.Default);
        return Parse(rtf, Path.GetFileNameWithoutExtension(path));
    }

    public WordDocument LoadFromString(string rtf, string? title = null) => Parse(rtf, title);

    public void Save(WordDocument model, string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        var rtf = Build(model);
        // UTF-8 BOM helps some editors; classic RTF is ANSI + \u escapes.
        File.WriteAllText(path, rtf, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    public string Build(WordDocument model)
    {
        var sb = new StringBuilder();
        sb.Append(@"{\rtf1\ansi\ansicpg65001\deff0");
        sb.Append(@"{\fonttbl{\f0\fnil\fcharset134 Microsoft JhengHei;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Arial;}}");
        sb.Append(@"{\colortbl;\red0\green0\blue0;}");
        sb.AppendLine();

        var fonts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["Microsoft JhengHei"] = 0,
            ["Calibri"] = 1,
            ["Arial"] = 2,
        };

        foreach (var block in model.Blocks)
        {
            switch (block)
            {
                case ParagraphBlock p:
                    AppendParagraph(sb, p, fonts);
                    break;
                case TableBlock t:
                    AppendTable(sb, t, fonts);
                    break;
                case ImageBlock img:
                    AppendImage(sb, img);
                    break;
            }
        }

        if (model.Blocks.Count == 0)
            sb.Append(@"\pard\par").AppendLine();

        sb.Append('}');
        return sb.ToString();
    }

    private static void AppendParagraph(StringBuilder sb, ParagraphBlock p, Dictionary<string, int> fonts)
    {
        sb.Append(@"\pard");
        sb.Append(p.Alignment switch
        {
            ParagraphAlignmentKind.Center => @"\qc",
            ParagraphAlignmentKind.Right => @"\qr",
            ParagraphAlignmentKind.Justify => @"\qj",
            _ => @"\ql",
        });

        if (p.ListKind == ListKind.Bullet)
            sb.Append(@"\li360 ").Append(EscapeText("• ")).Append(' ');
        else if (p.ListKind == ListKind.Numbered)
            sb.Append(@"\li360 ");

        var fi = ResolveFont(p.FontFamily, fonts);
        var halfPoints = Math.Clamp((int)Math.Round(p.FontSize * 2), 12, 144);
        sb.Append(@"\f").Append(fi).Append(@"\fs").Append(halfPoints);
        if (p.IsBold || p.Style != ParagraphStyleKind.Normal) sb.Append(@"\b");
        if (p.IsItalic) sb.Append(@"\i");
        if (p.IsUnderline) sb.Append(@"\ul");
        sb.Append(' ');
        sb.Append(EscapeText(p.Text ?? string.Empty));
        sb.Append(@"\b0\i0\ulnone\par").AppendLine();
    }

    private static void AppendTable(StringBuilder sb, TableBlock table, Dictionary<string, int> fonts)
    {
        table.EnsureRectangular();
        var cols = Math.Max(1, table.ColumnCount);
        var colWidth = 9000 / cols;

        foreach (var row in table.Rows)
        {
            sb.Append(@"\trowd\trgaph108\trleft0");
            for (var c = 0; c < cols; c++)
                sb.Append(@"\cellx").Append((c + 1) * colWidth);
            sb.AppendLine();

            for (var c = 0; c < cols; c++)
            {
                var cell = c < row.Cells.Count ? row.Cells[c] : new TableCellBlock();
                sb.Append(@"\pard\intbl");
                sb.Append(cell.Alignment switch
                {
                    ParagraphAlignmentKind.Center => @"\qc",
                    ParagraphAlignmentKind.Right => @"\qr",
                    ParagraphAlignmentKind.Justify => @"\qj",
                    _ => @"\ql",
                });
                var fi = ResolveFont(cell.FontFamily, fonts);
                var halfPoints = Math.Clamp((int)Math.Round(cell.FontSize * 2), 12, 144);
                sb.Append(@"\f").Append(fi).Append(@"\fs").Append(halfPoints);
                if (cell.IsBold) sb.Append(@"\b");
                if (cell.IsItalic) sb.Append(@"\i");
                if (cell.IsUnderline) sb.Append(@"\ul");
                sb.Append(' ');
                sb.Append(EscapeText(cell.Text ?? string.Empty));
                sb.Append(@"\b0\i0\ulnone\cell").AppendLine();
            }

            sb.Append(@"\row").AppendLine();
        }

        sb.Append(@"\pard\par").AppendLine();
    }

    private static void AppendImage(StringBuilder sb, ImageBlock img)
    {
        if (img.ImageBytes is not { Length: > 0 })
            return;

        var isJpeg = (img.ContentType ?? "").Contains("jpeg", StringComparison.OrdinalIgnoreCase)
                     || (img.ContentType ?? "").Contains("jpg", StringComparison.OrdinalIgnoreCase)
                     || (img.FileName?.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) == true)
                     || (img.FileName?.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) == true);

        var widthTwips = (int)Math.Clamp(img.DisplayWidth / 96.0 * 1440.0, 200, 10000);
        var heightDip = img.DisplayHeight > 0 ? img.DisplayHeight : img.ResolvedHeight;
        var heightTwips = (int)Math.Clamp(heightDip / 96.0 * 1440.0, 200, 14000);

        sb.Append(@"\pard{\pict");
        sb.Append(isJpeg ? @"\jpegblip" : @"\pngblip");
        sb.Append(@"\picwgoal").Append(widthTwips);
        sb.Append(@"\pichgoal").Append(heightTwips);
        sb.Append(' ');
        foreach (var b in img.ImageBytes)
            sb.Append(b.ToString("x2", CultureInfo.InvariantCulture));
        sb.Append(@"}\par").AppendLine();
    }

    private static int ResolveFont(string? family, Dictionary<string, int> fonts)
    {
        if (string.IsNullOrWhiteSpace(family))
            return 0;
        if (fonts.TryGetValue(family, out var id))
            return id;
        // closest default
        if (family.Contains("Arial", StringComparison.OrdinalIgnoreCase)) return 2;
        if (family.Contains("Calibri", StringComparison.OrdinalIgnoreCase)) return 1;
        return 0;
    }

    private static string EscapeText(string text)
    {
        var sb = new StringBuilder(text.Length * 2);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '\\': sb.Append(@"\\"); break;
                case '{': sb.Append(@"\{"); break;
                case '}': sb.Append(@"\}"); break;
                case '\r': break;
                case '\n': sb.Append(@"\line "); break;
                case '\t': sb.Append(@"\tab "); break;
                default:
                    if (ch <= 0x7F)
                        sb.Append(ch);
                    else
                    {
                        // signed 16-bit Unicode for RTF \uN?
                        short code = unchecked((short)ch);
                        sb.Append(@"\u").Append(code.ToString(CultureInfo.InvariantCulture)).Append('?');
                    }
                    break;
            }
        }

        return sb.ToString();
    }

    // ─── Parse (tolerant, model-oriented) ───────────────────

    private static WordDocument Parse(string rtf, string? title)
    {
        var doc = new WordDocument { Title = title ?? "文件" };
        if (string.IsNullOrWhiteSpace(rtf) || !rtf.Contains(@"\rtf", StringComparison.OrdinalIgnoreCase))
        {
            doc.Blocks.Add(new ParagraphBlock());
            return doc;
        }

        // Strip font/color tables groups roughly by brace matching for top-level groups after header.
        var body = StripHeaderTables(rtf);

        // Split into paragraphs on \par (not inside \pict hex roughly)
        var parts = SplitParagraphs(body);
        TableBlock? currentTable = null;

        foreach (var part in parts)
        {
            if (string.IsNullOrWhiteSpace(part))
                continue;

            if (part.Contains(@"\trowd", StringComparison.Ordinal))
            {
                currentTable ??= new TableBlock();
                var row = ParseTableRow(part);
                if (row.Cells.Count > 0)
                    currentTable.Rows.Add(row);
                continue;
            }

            if (currentTable is not null)
            {
                doc.Blocks.Add(currentTable);
                currentTable = null;
            }

            if (part.Contains(@"\pict", StringComparison.Ordinal))
            {
                var img = TryParseImage(part);
                if (img is not null)
                    doc.Blocks.Add(img);
                // also keep any leading text
                var textOnly = Regex.Replace(part, @"\{\\pict[\s\S]*?\}", "", RegexOptions.IgnoreCase);
                var p = ParseParagraph(textOnly);
                if (!string.IsNullOrWhiteSpace(p.Text))
                    doc.Blocks.Add(p);
                continue;
            }

            doc.Blocks.Add(ParseParagraph(part));
        }

        if (currentTable is not null)
            doc.Blocks.Add(currentTable);

        if (doc.Blocks.Count == 0)
            doc.Blocks.Add(new ParagraphBlock());

        return doc;
    }

    private static string StripHeaderTables(string rtf)
    {
        // Find first \pard or content after fonttbl/colortbl
        var idx = rtf.IndexOf(@"\pard", StringComparison.OrdinalIgnoreCase);
        if (idx < 0)
            idx = rtf.IndexOf(@"\sectd", StringComparison.OrdinalIgnoreCase);
        if (idx < 0)
        {
            // skip leading {\rtf1...{\fonttbl...}{\colortbl...}
            var brace = 0;
            var i = 0;
            var skippedGroups = 0;
            while (i < rtf.Length && skippedGroups < 3)
            {
                if (rtf[i] == '{')
                {
                    brace++;
                    i++;
                    continue;
                }

                if (rtf[i] == '}')
                {
                    brace--;
                    i++;
                    if (brace == 1)
                        skippedGroups++;
                    continue;
                }

                i++;
            }

            return i < rtf.Length ? rtf[i..] : rtf;
        }

        return rtf[idx..];
    }

    private static List<string> SplitParagraphs(string body)
    {
        // Split on \par not followed by d (avoid \pard) — use \par with word boundary
        var list = new List<string>();
        var sb = new StringBuilder();
        for (var i = 0; i < body.Length; i++)
        {
            if (body[i] == '\\' && i + 3 < body.Length
                && body[i + 1] == 'p' && body[i + 2] == 'a' && body[i + 3] == 'r'
                && (i + 4 >= body.Length || !char.IsLetter(body[i + 4])))
            {
                list.Add(sb.ToString());
                sb.Clear();
                i += 3;
                if (i + 1 < body.Length && body[i + 1] == ' ')
                    i++;
                continue;
            }

            sb.Append(body[i]);
        }

        if (sb.Length > 0)
            list.Add(sb.ToString());
        return list;
    }

    private static TableRowBlock ParseTableRow(string part)
    {
        var row = new TableRowBlock();
        var cells = Regex.Split(part, @"\\cell(?![a-zA-Z])");
        foreach (var cell in cells)
        {
            if (cell.Contains(@"\trowd", StringComparison.Ordinal) && !cell.Contains(@"\intbl", StringComparison.Ordinal))
                continue;
            if (string.IsNullOrWhiteSpace(cell) || cell.Trim() == @"\row")
                continue;
            var p = ParseParagraph(cell);
            // skip pure control remnants
            if (string.IsNullOrEmpty(p.Text) && !cell.Contains(@"\intbl", StringComparison.Ordinal))
                continue;
            row.Cells.Add(new TableCellBlock
            {
                Text = p.Text,
                IsBold = p.IsBold,
                IsItalic = p.IsItalic,
                IsUnderline = p.IsUnderline,
                FontSize = p.FontSize,
                FontFamily = p.FontFamily,
                Alignment = p.Alignment,
            });
        }

        return row;
    }

    private static ParagraphBlock ParseParagraph(string part)
    {
        var alignment = ParagraphAlignmentKind.Left;
        if (part.Contains(@"\qc", StringComparison.Ordinal)) alignment = ParagraphAlignmentKind.Center;
        else if (part.Contains(@"\qr", StringComparison.Ordinal)) alignment = ParagraphAlignmentKind.Right;
        else if (part.Contains(@"\qj", StringComparison.Ordinal)) alignment = ParagraphAlignmentKind.Justify;

        var bold = Regex.IsMatch(part, @"\\b(?![0a-zA-Z])");
        var italic = Regex.IsMatch(part, @"\\i(?![0a-zA-Z])");
        var underline = Regex.IsMatch(part, @"\\ul(?![a-zA-Z])");

        double fontSize = 14;
        var fs = Regex.Match(part, @"\\fs(\d+)");
        if (fs.Success && int.TryParse(fs.Groups[1].Value, out var half))
            fontSize = half / 2.0;

        var listKind = ListKind.None;
        if (part.Contains('•') || part.Contains(@"\bullet", StringComparison.Ordinal))
            listKind = ListKind.Bullet;

        var text = DecodeRtfText(part);
        // strip leading bullet we injected
        if (text.StartsWith("• "))
        {
            text = text[2..];
            listKind = ListKind.Bullet;
        }

        var style = ParagraphStyleKind.Normal;
        if (fontSize >= 26 && bold) style = ParagraphStyleKind.Heading1;
        else if (fontSize >= 20 && bold) style = ParagraphStyleKind.Heading2;
        else if (fontSize >= 16 && bold && fontSize < 20) style = ParagraphStyleKind.Heading3;

        return new ParagraphBlock
        {
            Text = text,
            Alignment = alignment,
            IsBold = bold,
            IsItalic = italic,
            IsUnderline = underline,
            FontSize = fontSize,
            FontFamily = "Microsoft JhengHei",
            ListKind = listKind,
            Style = style,
        };
    }

    private static string DecodeRtfText(string part)
    {
        var sb = new StringBuilder();
        // Remove pict groups first
        part = Regex.Replace(part, @"\{\\pict[\s\S]*?\}", "", RegexOptions.IgnoreCase);

        for (var i = 0; i < part.Length; i++)
        {
            var ch = part[i];
            if (ch == '\\')
            {
                if (i + 1 >= part.Length)
                    break;
                var next = part[i + 1];
                if (next is '\\' or '{' or '}')
                {
                    sb.Append(next);
                    i++;
                    continue;
                }

                // Unicode char: \uN or \u-N (must be followed by digit / minus+digit; not \ul, \ulnone, …)
                if (next == 'u'
                    && i + 2 < part.Length
                    && (char.IsDigit(part[i + 2]) || (part[i + 2] == '-' && i + 3 < part.Length && char.IsDigit(part[i + 3]))))
                {
                    i += 2;
                    var sign = 1;
                    if (i < part.Length && part[i] == '-')
                    {
                        sign = -1;
                        i++;
                    }

                    var num = 0;
                    var any = false;
                    while (i < part.Length && char.IsDigit(part[i]))
                    {
                        any = true;
                        num = num * 10 + (part[i] - '0');
                        i++;
                    }

                    if (any)
                    {
                        var code = sign < 0 ? -num : num;
                        if (code < 0)
                            code += 65536;
                        if (code is > 0 and < 0x110000)
                            sb.Append(char.ConvertFromUtf32(Math.Min(code, 0x10FFFF)));
                    }

                    // optional ANSI fallback char after \uN
                    if (i < part.Length && part[i] is not ('\\' or ' ' or '\r' or '\n' or '{' or '}'))
                        i++; // skip fallback
                    i--; // loop will i++
                    continue;
                }

                if (next == '\'')
                {
                    // hex byte \'hh
                    if (i + 3 < part.Length
                        && Uri.IsHexDigit(part[i + 2])
                        && Uri.IsHexDigit(part[i + 3]))
                    {
                        var hex = part.Substring(i + 2, 2);
                        var b = Convert.ToByte(hex, 16);
                        sb.Append(Encoding.GetEncoding(1252).GetString(new[] { b }));
                        i += 3;
                        continue;
                    }
                }

                // control word: skip letters and optional digits
                i++;
                while (i < part.Length && char.IsLetter(part[i]))
                    i++;
                if (i < part.Length && part[i] == '-')
                    i++;
                while (i < part.Length && char.IsDigit(part[i]))
                    i++;
                if (i < part.Length && part[i] == ' ')
                {
                    // space delimiter consumed
                }
                else
                    i--;

                // map a few useful controls
                continue;
            }

            if (ch is '{' or '}')
                continue;
            if (ch is '\r' or '\n')
                continue;
            sb.Append(ch);
        }

        return sb.ToString().Trim();
    }

    private static ImageBlock? TryParseImage(string part)
    {
        var m = Regex.Match(part, @"\\pict[^}]*?(?:\\pngblip|\\jpegblip)[^}]*?([0-9a-fA-F\s]+)(?=\})", RegexOptions.IgnoreCase);
        if (!m.Success)
        {
            // fallback: after blip keyword
            m = Regex.Match(part, @"\\(?:png|jpeg)blip[^\\]*?([0-9a-fA-F\r\n\s]{32,})", RegexOptions.IgnoreCase);
        }

        if (!m.Success)
            return null;

        var hex = Regex.Replace(m.Groups[1].Value, @"\s+", "");
        if (hex.Length < 16 || hex.Length % 2 != 0)
            return null;

        try
        {
            var bytes = Convert.FromHexString(hex);
            var isJpeg = part.Contains(@"\jpegblip", StringComparison.OrdinalIgnoreCase);
            double w = 400, h = 0;
            var wg = Regex.Match(part, @"\\picwgoal(\d+)");
            var hg = Regex.Match(part, @"\\pichgoal(\d+)");
            if (wg.Success && int.TryParse(wg.Groups[1].Value, out var tw))
                w = Math.Clamp(tw / 1440.0 * 96.0, 40, 720);
            if (hg.Success && int.TryParse(hg.Groups[1].Value, out var th))
                h = Math.Max(0, th / 1440.0 * 96.0);

            return new ImageBlock
            {
                ImageBytes = bytes,
                ContentType = isJpeg ? "image/jpeg" : "image/png",
                FileName = isJpeg ? "image.jpg" : "image.png",
                DisplayWidth = w,
                DisplayHeight = h,
            };
        }
        catch
        {
            return null;
        }
    }
}
