using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;
using DocxAvalonia.Models;

namespace DocxAvalonia.Converters;

public static class EditorConverters
{
    public static readonly IValueConverter FontWeightFromBool =
        new FuncValueConverter<bool, FontWeight>(b => b ? FontWeight.Bold : FontWeight.Normal);

    public static readonly IValueConverter FontStyleFromBool =
        new FuncValueConverter<bool, FontStyle>(b => b ? FontStyle.Italic : FontStyle.Normal);

    public static readonly IValueConverter TextDecorationsFromBool =
        new FuncValueConverter<bool, TextDecorationCollection?>(b => b ? TextDecorations.Underline : null);

    public static readonly IValueConverter AlignmentFromKind =
        new FuncValueConverter<ParagraphAlignmentKind, TextAlignment>(a => a switch
        {
            ParagraphAlignmentKind.Center => TextAlignment.Center,
            ParagraphAlignmentKind.Right => TextAlignment.Right,
            ParagraphAlignmentKind.Justify => TextAlignment.Justify,
            _ => TextAlignment.Left,
        });

    public static readonly IValueConverter ListPrefix =
        new FuncValueConverter<ListKind, string>(k => k switch
        {
            ListKind.Bullet => "•  ",
            ListKind.Numbered => "1. ",
            _ => string.Empty,
        });

    public static readonly IValueConverter IsParagraph =
        new FuncValueConverter<DocumentBlock?, bool>(b => b is ParagraphBlock);

    public static readonly IValueConverter StyleLabel =
        new FuncValueConverter<ParagraphStyleKind, string>(s => s switch
        {
            ParagraphStyleKind.Heading1 => "標題 1",
            ParagraphStyleKind.Heading2 => "標題 2",
            ParagraphStyleKind.Heading3 => "標題 3",
            _ => "內文",
        });

    public static readonly IValueConverter FontFamilyFromName =
        new FuncValueConverter<string?, FontFamily>(name =>
            string.IsNullOrWhiteSpace(name)
                ? new FontFamily("Microsoft JhengHei, Microsoft YaHei, Segoe UI, sans-serif")
                : new FontFamily(name));

    /// <summary>True when bound string equals ConverterParameter (ribbon tab, UI theme, etc.).</summary>
    public static readonly IValueConverter IsRibbonTab = new StringEqualsConverter();

    /// <summary>Alias for theme / layout string equality checks.</summary>
    public static readonly IValueConverter IsUiTheme = IsRibbonTab;

    public static readonly IValueConverter UiThemeLabel =
        new FuncValueConverter<string?, string>(id => id switch
        {
            "Word" => "Microsoft Word",
            "LibreOffice" => "LibreOffice Writer",
            "GoogleDocs" => "Google 文件",
            "Zoho" => "Zoho Writer",
            "Wps" => "WPS Writer",
            "FreeOffice" => "FreeOffice TextMaker",
            _ => id ?? string.Empty,
        });

    private sealed class StringEqualsConverter : IValueConverter
    {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
            => string.Equals(value?.ToString(), parameter?.ToString(), StringComparison.OrdinalIgnoreCase);

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
