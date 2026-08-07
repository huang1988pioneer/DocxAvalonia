using System.IO;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Binary.WordprocessingMLMapping;
using DocxAvalonia.Models;
using BinaryDoc = DocSharp.Binary.DocFileFormat.WordDocument;
using BinaryOpenXml = DocSharp.Binary.OpenXmlLib;
using BinaryWord = DocSharp.Binary.OpenXmlLib.WordprocessingML;

namespace DocxAvalonia.Services;

/// <summary>
/// Read Word 97-2003 binary .doc (via DocSharp → DOCX) and read/write RTF-based .doc
/// (our export format; Word / WPS / LibreOffice open it).
/// </summary>
public sealed class DocBinaryDocumentService
{
    private static readonly byte[] OleMagic = [0xD0, 0xCF, 0x11, 0xE0];
    private readonly RtfDocumentService _rtf = new();

    public WordDocument Load(string path, DocxDocumentService docx)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("找不到指定的文件。", path);

        var header = new byte[8];
        int read;
        using (var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            read = fs.Read(header, 0, header.Length);

        // RTF saved as .doc (our export path and some other editors)
        if (IsRtf(header, read))
            return _rtf.Load(path);

        // OLE compound document (classic Word 97-2003 .doc)
        if (IsOle(header, read))
            return LoadBinaryDocAsDocx(path, docx);

        // Some misnamed OOXML packages
        if (read >= 2 && header[0] == (byte)'P' && header[1] == (byte)'K')
            return docx.Load(path);

        throw new InvalidOperationException("無法辨識的 .doc 檔案格式（需為 Word 97-2003 二進位或 RTF）。");
    }

    public void Save(WordDocument model, string path, DocxDocumentService docx)
    {
        // True binary Word 97 write is not available in free pure-managed stacks.
        // RTF with .doc extension opens in Word / WPS / LibreOffice and round-trips here.
        try
        {
            _rtf.Save(model, path);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"儲存 .doc 失敗：{ex.Message}", ex);
        }
    }

    private static WordDocument LoadBinaryDocAsDocx(string path, DocxDocumentService docx)
    {
        var tempDocx = Path.Combine(Path.GetTempPath(), $"DocxAvalonia-{Guid.NewGuid():N}.docx");
        try
        {
            using (var reader = new StructuredStorageReader(path))
            {
                var binaryDoc = new BinaryDoc(reader);
                using var openXmlDoc = BinaryWord.WordprocessingDocument.Create(
                    tempDocx,
                    BinaryOpenXml.WordprocessingDocumentType.Document);
                Converter.Convert(binaryDoc, openXmlDoc);
            }

            return docx.Load(tempDocx);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"讀取 .doc 失敗：{ex.Message}", ex);
        }
        finally
        {
            TryDelete(tempDocx);
        }
    }

    private static bool IsOle(byte[] header, int length) =>
        length >= 4
        && header[0] == OleMagic[0]
        && header[1] == OleMagic[1]
        && header[2] == OleMagic[2]
        && header[3] == OleMagic[3];

    private static bool IsRtf(byte[] header, int length)
    {
        if (length < 5)
            return false;
        return header[0] == (byte)'{'
               && header[1] == (byte)'\\'
               && header[2] == (byte)'r'
               && header[3] == (byte)'t'
               && header[4] == (byte)'f';
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // ignore temp cleanup failures
        }
    }
}
