using System.IO;
using System.Text;
using DocxAvalonia.Models;

namespace DocxAvalonia.Services;

/// <summary>
/// Unified load/save for .docx, .doc (Word 97-2003 / RTF-compatible), and .odt.
/// </summary>
public sealed class DocumentFormatService
{
    public static readonly string[] SupportedExtensions = [".docx", ".doc", ".odt"];

    private readonly DocxDocumentService _docx = new();
    private readonly OdtDocumentService _odt = new();
    private readonly DocBinaryDocumentService _doc = new();

    static DocumentFormatService()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public static bool IsSupportedPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return false;
        var ext = Path.GetExtension(path);
        return SupportedExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase));
    }

    public static string NormalizeSavePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路徑不可為空。", nameof(path));

        var ext = Path.GetExtension(path);
        if (string.IsNullOrEmpty(ext))
            return path + ".docx";

        if (!IsSupportedPath(path))
            throw new InvalidOperationException($"不支援的格式「{ext}」。請使用 .docx、.doc 或 .odt。");

        return path;
    }

    public WordDocument Load(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("找不到指定的文件。", path);

        var ext = Path.GetExtension(path);
        return ext.ToLowerInvariant() switch
        {
            ".docx" => _docx.Load(path),
            ".doc" => _doc.Load(path, _docx),
            ".odt" => _odt.Load(path),
            _ => throw new InvalidOperationException($"僅支援 .docx、.doc、.odt 格式（目前為 {ext}）。"),
        };
    }

    public void Save(WordDocument model, string path)
    {
        path = NormalizeSavePath(path);
        var dir = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        var ext = Path.GetExtension(path).ToLowerInvariant();
        switch (ext)
        {
            case ".docx":
                _docx.Save(model, path);
                break;
            case ".doc":
                _doc.Save(model, path, _docx);
                break;
            case ".odt":
                _odt.Save(model, path);
                break;
            default:
                throw new InvalidOperationException($"僅支援 .docx、.doc、.odt 格式（目前為 {ext}）。");
        }
    }
}
