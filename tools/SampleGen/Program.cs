using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

var outputPath = args.Length > 0
    ? args[0]
    : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "Samples", "demo.docx"));

// Prefer repo Samples folder when running from tools/SampleGen
var repoSample = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "Samples", "demo.docx"));
if (args.Length == 0 && Directory.Exists(Path.GetDirectoryName(repoSample)!))
    outputPath = repoSample;

Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);

using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
var main = doc.AddMainDocumentPart();
main.Document = new Document();
var body = main.Document.AppendChild(new Body());

var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
numberingPart.Numbering = new Numbering(
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

var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
stylesPart.Styles = new Styles(
    new Style(
        new StyleName { Val = "heading 1" },
        new BasedOn { Val = "Normal" },
        new UIPriority { Val = 9 },
        new PrimaryStyle(),
        new StyleRunProperties(new Bold(), new FontSize { Val = "32" })
    ) { Type = StyleValues.Paragraph, StyleId = "Heading1" },
    new Style(
        new StyleName { Val = "heading 2" },
        new BasedOn { Val = "Normal" },
        new UIPriority { Val = 9 },
        new PrimaryStyle(),
        new StyleRunProperties(new Bold(), new FontSize { Val = "26" })
    ) { Type = StyleValues.Paragraph, StyleId = "Heading2" }
);

body.AppendChild(CreateParagraph("DocxAvalonia 預覽範例", "Heading1", bold: true, size: "36"));
body.AppendChild(CreateParagraph("這是一份用來驗證 Avalonia .docx 預覽的範例文件。", null, false, null, "0000AA"));
body.AppendChild(CreateParagraph("功能展示", "Heading2", bold: true));
body.AppendChild(CreateParagraph("支援粗體、斜體、底線與顏色。", null, bold: true));
body.AppendChild(CreateFormattedParagraph());
body.AppendChild(CreateListItem("項目符號清單第一點", 1));
body.AppendChild(CreateListItem("項目符號清單第二點", 1));
body.AppendChild(CreateListItem("編號清單項目一", 2));
body.AppendChild(CreateListItem("編號清單項目二", 2));
body.AppendChild(CreateParagraph("資料表格", "Heading2", bold: true));
body.AppendChild(CreateSampleTable());
body.AppendChild(CreateParagraph("文件結束。謝謝使用 DocxAvalonia！", null));

main.Document.Save();
Console.WriteLine("Created: " + Path.GetFullPath(outputPath));

static Paragraph CreateParagraph(string text, string? style, bool bold = false, string? size = null, string? color = null)
{
    var pPr = new ParagraphProperties();
    if (style is not null)
        pPr.AppendChild(new ParagraphStyleId { Val = style });

    var rPr = new RunProperties();
    if (bold) rPr.AppendChild(new Bold());
    if (size is not null) rPr.AppendChild(new FontSize { Val = size });
    if (color is not null) rPr.AppendChild(new Color { Val = color });
    rPr.AppendChild(new RunFonts
    {
        Ascii = "Microsoft JhengHei",
        EastAsia = "Microsoft JhengHei",
        HighAnsi = "Calibri",
    });

    return new Paragraph(pPr, new Run(rPr, new Text(text)));
}

static Paragraph CreateFormattedParagraph()
{
    var p = new Paragraph();
    p.AppendChild(new Run(new RunProperties(new Bold()), new Text("粗體")));
    p.AppendChild(new Run(new Text("、")));
    p.AppendChild(new Run(new RunProperties(new Italic()), new Text("斜體")));
    p.AppendChild(new Run(new Text("、")));
    p.AppendChild(new Run(new RunProperties(new Underline { Val = UnderlineValues.Single }), new Text("底線")));
    p.AppendChild(new Run(new Text(" 與 ")));
    p.AppendChild(new Run(new RunProperties(new Color { Val = "C00000" }, new Bold()), new Text("紅色文字")));
    p.AppendChild(new Run(new Text("。")));
    return p;
}

static Paragraph CreateListItem(string text, int numId)
{
    var pPr = new ParagraphProperties(
        new NumberingProperties(
            new NumberingLevelReference { Val = 0 },
            new NumberingId { Val = numId }
        )
    );
    return new Paragraph(pPr, new Run(new Text(text)));
}

static Table CreateSampleTable()
{
    var table = new Table(
        new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
            )
        )
    );

    table.AppendChild(CreateRow(("功能", "狀態", "備註"), header: true));
    table.AppendChild(CreateRow(("段落預覽", "✓", "含格式"), header: false));
    table.AppendChild(CreateRow(("表格預覽", "✓", "簡易格線"), header: false));
    table.AppendChild(CreateRow(("清單預覽", "✓", "符號/編號"), header: false));
    return table;
}

static TableRow CreateRow((string A, string B, string C) cells, bool header)
{
    var row = new TableRow();
    foreach (var cellText in new[] { cells.A, cells.B, cells.C })
    {
        var rPr = new RunProperties();
        if (header) rPr.AppendChild(new Bold());
        var cell = new TableCell(
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "1666" }),
            new Paragraph(new Run(rPr, new Text(cellText)))
        );
        row.AppendChild(cell);
    }
    return row;
}
