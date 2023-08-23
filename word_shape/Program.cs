using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main(string[] args)
    {
        string currentDirectory = Directory.GetCurrentDirectory();
        string[] wordFiles = Directory.GetFiles(currentDirectory, "*.docx");

        foreach (string filePath in wordFiles)
        {
            ApplyFormatting(filePath);
        }

        Console.WriteLine("Formatting applied to all Word files in the current directory.");
    }

    static void ApplyFormatting(string filePath)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            // 新しいフォント設定を作成
            RunFonts runFonts = new RunFonts() { Ascii = "游明朝", HighAnsi = "游明朝", EastAsia = "游明朝" };
            FontSize fontSize = new FontSize() { Val = "21" };  // 10.5ポイント (1/2 ポイント単位)

            // 新しい段落設定を作成
            SpacingBetweenLines spacing = new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Line = "276", Before = "0", After = "0" };  // 1.15倍の行間 (1/20 ポイント単位)

            // 新しいページ余白設定を作成
            PageMargin pageMargin = new PageMargin() { Top = 11340, Bottom = 11340, Left = 22680, Right = 11340 };  // 15mm = 1417 ポイント (1/20 ポイント単位)

            // ドキュメント内のすべてのランを取得し、フォント設定を適用
            foreach (var run in doc.MainDocumentPart.Document.Descendants<Run>())
            {
                run.RunProperties = new RunProperties(runFonts, fontSize);
            }

            // ドキュメント内のすべての段落を取得し、段落設定を適用
            foreach (var paragraph in doc.MainDocumentPart.Document.Descendants<Paragraph>())
            {
                paragraph.Append(spacing);
            }

            // ドキュメント内のすべてのセクションを取得し、ページ余白設定を適用
            foreach (var section in doc.MainDocumentPart.Document.Body.Elements<SectionProperties>())
            {
                section.Append(pageMargin);
            }
            // 変更を保存
            doc.MainDocumentPart.Document.Save();
        }
    }
}
