using Aspose.Words;
using Microsoft.AspNetCore.Mvc;
using Paragraph = Aspose.Words.Paragraph;
using Document = Aspose.Words.Document;
using DocumentFormat.OpenXml.Drawing;
using Run = Aspose.Words.Run;

namespace WordStyleService.Controllers;

[ApiController]
[Route("api/[controller]")]
public class WordController : ControllerBase
{
    [HttpPost("analyze")]
    public IActionResult AnalyzeDocument([FromBody] DocumentRequest request)
    {
        var docBytes = Convert.FromBase64String(request.Base64Doc);
        var paraStyles = new Dictionary<string, string>();
        var stylesInline = new List<string>();

        // Use 'using' statement to properly dispose of the MemoryStream
        using var stream = new MemoryStream(docBytes);
        var doc = new Document(stream);

        // Approach 2: Simplified with separate methods (cleaner and more readable)
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>();

        // Process paragraph styles
        ProcessParagraphStyles(paragraphs, paraStyles);

        // Process inline styles
        ProcessInlineStyles(paragraphs, stylesInline);

        var result = new
        {
            ParagraphStyles = paraStyles,
            InlineStyles = stylesInline
        };
        doc.AcceptAllRevisions();
        return Ok(result);
    }

    [HttpPost("apply-style")]
    public IActionResult ApplyStyle([FromBody] StyleRequest request)
    {
        var docBytes = Convert.FromBase64String(request.Base64Doc);

        using var inputStream = new MemoryStream(docBytes);
        var doc = new Document(inputStream);

        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            var text = para.GetText().Trim();
            if (request.Styles.ContainsKey(text))
            {
                if (text == "Development Test")
                {
                    Console.WriteLine("Style to Development", para.ParagraphFormat.StyleName);
                }
                para.ParagraphFormat.StyleName = request.Styles[text];
            }
        }

        using var outputStream = new MemoryStream();
        doc.Save(outputStream, SaveFormat.Docx);
        var result = Convert.ToBase64String(outputStream.ToArray());
        return Ok(result);
    }

    [HttpPost("compare-and-apply")]
    public IActionResult CompareAndApply([FromBody] CompareRequest request)
    {
        var sourceStyles = ExtractStyles(request.SourceBase64);
        using var sourceStream = new MemoryStream(Convert.FromBase64String(request.SourceBase64));
        using var targetStream = new MemoryStream(Convert.FromBase64String(request.TargetBase64));

        // compare docs
        var docA = new Document(sourceStream);
        var docB = new Document(targetStream);

        docA.AcceptAllRevisions();
        docB.AcceptAllRevisions();
        docB.Compare(docA, "Ivan", DateTime.Now);
        Console.WriteLine(docB.Revisions.Count == 0 ? "Documents are equal" : "Documents are different");

        using var ms = new MemoryStream();
        docB.Save(ms, SaveFormat.Docx);
        return Ok(Convert.ToBase64String(ms.ToArray()));
    }

    private Dictionary<string, string> ExtractStyles(string base64Doc)
    {

        var styles = new Dictionary<string, string>();
        using var stream = new MemoryStream(Convert.FromBase64String(base64Doc));
        var doc = new Document(stream);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            var text = para.GetText().Trim();
            var styleName = para.ParagraphFormat?.Style?.Name ?? "Normal";
            if (!string.IsNullOrEmpty(text))
            {
                styles[text] = styleName;
            }
        }
        return styles;
    }

    // Helper method to process paragraph styles
    private void ProcessParagraphStyles(IEnumerable<Paragraph> paragraphs, Dictionary<string, string> paraStyles)
    {
        paragraphs
            .Where(para => !string.IsNullOrEmpty(para.GetText().Trim()))
            .ToList()
            .ForEach(para =>
            {
                var text = para.GetText().Trim();
                var style = para.ParagraphFormat.StyleName;

                if (text == "Development")
                {
                    Console.WriteLine("Style to Development", style);
                }
                paraStyles[text] = style;
            });
    }

    // Helper method to process inline styles
    private void ProcessInlineStyles(IEnumerable<Paragraph> paragraphs, List<string> stylesInline)
    {
        int position = 0;

        // Approach 2a: Using Parallel.ForEach for better performance (if you have many paragraphs)
        foreach (var para in paragraphs)
        {
            var baseFont = para.ParagraphFormat.Style.Font;
            var text = para.GetText().Trim();
            var style = para.ParagraphFormat.StyleName;

            // Process runs within this paragraph
            para.Runs.Cast<Run>()
                .Where(run => HasFontDifferences(run.Font, baseFont) && run.Font.StyleName == "Default Paragraph Font")
                .ToList()
                .ForEach(run =>
                {
                    Console.WriteLine("old style-->" + style);
                    Console.WriteLine(baseFont.Name);
                    var changes = $"{run.Font.Name} {run.Font.Size} {run.Font.Bold} {run.Font.Italic} {run.Font.Underline} {run.Font.StrikeThrough} {run.Font.Color}";
                    Console.WriteLine("new style-->" + run.Font.StyleName + "  " + changes + " for text: " + text);
                    stylesInline.Add($"Font: {run.Font.Name} Position: {position} Text: {run.Text}");
                });

            position += para.GetText().Length;
        }
    }

    // Helper method to check font differences
    private bool HasFontDifferences(Font runFont, Font baseFont)
    {
        return runFont.Name != baseFont.Name ||
               runFont.Size != baseFont.Size ||
               runFont.Bold != baseFont.Bold ||
               runFont.Italic != baseFont.Italic ||
               runFont.Underline != baseFont.Underline ||
               runFont.StrikeThrough != baseFont.StrikeThrough ||
               runFont.Color != baseFont.Color;
    }
}
public class CompareRequest
{
    public required string SourceBase64 { get; set; }
    public required string TargetBase64 { get; set; }
}
public class DocumentRequest
{
    public required string Base64Doc { get; set; }
}

public class StyleRequest
{
    public required string Base64Doc { get; set; }
    public required Dictionary<string, string> Styles { get; set; }
}