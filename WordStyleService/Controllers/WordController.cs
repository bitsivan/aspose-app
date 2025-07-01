using Aspose.Words;
using Microsoft.AspNetCore.Mvc;
using Paragraph = Aspose.Words.Paragraph;
using Document = Aspose.Words.Document;

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
        var styles = new Dictionary<string, string>();
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            var style = para.ParagraphFormat.StyleName;
            var text = para.GetText().Trim();
            if (!string.IsNullOrEmpty(text))
            {
                styles[text] = style;
            }
        }
        return Ok(styles);
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
        var targetDocBytes = Convert.FromBase64String(request.TargetBase64);

        using var stream = new MemoryStream(targetDocBytes);
        var doc=new Document(stream);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (var paragraph in paragraphs)
        {
            var para = paragraph as Paragraph;
            var text = para?.GetText().Trim();
            if (para != null && !string.IsNullOrEmpty(text) && sourceStyles.ContainsKey(text))
            {
                para.ParagraphFormat.StyleName = sourceStyles[text];
            }
        }

        using var ms = new MemoryStream();
        doc.Save(ms, SaveFormat.Docx);
        return Ok(Convert.ToBase64String(ms.ToArray()));
    }

    private Dictionary<string, string> ExtractStyles(string base64Doc)
    {

        var styles = new Dictionary<string, string>();
        using var stream = new MemoryStream(Convert.FromBase64String(base64Doc));
        var doc=new Document(stream);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            var text = para.GetText().Trim();
            var styleName= para.ParagraphFormat?.Style?.Name ?? "Normal";
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