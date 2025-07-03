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
        var doc = new Document(new MemoryStream(docBytes));
        var paraStyles = new Dictionary<string, string>();
        var stylesInline = new List<string>();

        // int segmentStart = 0;
        int position = 0;
        // string currentStyle = null;
        //Iterate Paragraph Nodes
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            var style = para.ParagraphFormat.StyleName;
            Style baseStyle = para.ParagraphFormat.Style;
            var baseFont = para.ParagraphFormat.Style.Font;
            var text = para.GetText().Trim();

            if (text == "Development")
            {
                Console.WriteLine("Style to Development", para.ParagraphFormat.StyleName);
            }
            if (!string.IsNullOrEmpty(text))
            {
                paraStyles[text] = style;
            }

            foreach (Run run in para.Runs)
            {
                var RunText = run.Text;
                var runFont = run.Font;
                string styleRun = run.Font.StyleName;

                if ((runFont.Name != baseFont.Name ||
                   runFont.Size != baseFont.Size ||
                   runFont.Bold != baseFont.Bold ||
                   runFont.Italic != baseFont.Italic ||
                   runFont.Underline != baseFont.Underline ||
                   runFont.StrikeThrough != baseFont.StrikeThrough ||
                   runFont.Color != baseFont.Color) && styleRun == "Default Paragraph Font")
                {
                    Console.WriteLine("old style-->" + style);
                    Console.WriteLine(baseFont.Name);
                    var changes = $"{run.Font.Name} {run.Font.Size} {run.Font.Bold} {run.Font.Italic} {run.Font.Underline} {run.Font.StrikeThrough} {run.Font.Color}";
                    Console.WriteLine("new style-->" + styleRun + "  " + changes + " for text: " + text);
                    stylesInline.Add($"Font: {runFont.Name} Position: {position} Text: {RunText}");
                }
                position += RunText.Length;

            }
        }

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
        var doc = new Document(new MemoryStream(docBytes));
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

        using var ms = new MemoryStream();
        doc.Save(ms, SaveFormat.Docx);
        var result = Convert.ToBase64String(ms.ToArray());
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

}

public class CompareRequest
{
    public string SourceBase64 { get; set; }
    public string TargetBase64 { get; set; }
}
public class DocumentRequest
{
    public string Base64Doc { get; set; }
}

public class StyleRequest
{
    public string Base64Doc { get; set; }
    public Dictionary<string, string> Styles { get; set; }
}