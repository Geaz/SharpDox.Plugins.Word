using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using NotesFor.HtmlToOpenXml;
using System.Collections.Generic;

namespace SharpDox.Plugins.Word.OpenXml.Elements
{
    internal class RichText : BaseElement
    {
        public RichText(string content) : base(content) { }

        public override void AppendTo(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            var paragraphs = ConvertContentToParagraphs(mainDocumentPart);
            foreach (var paragraph in paragraphs)
            {
                openXmlNode.Append(paragraph);
            }
        }

        public override void InsertAfter(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            var paragraphs = ConvertContentToParagraphs(mainDocumentPart);
            var insertPoint = openXmlNode;
            foreach (var paragraph in paragraphs)
            {
                insertPoint.InsertAfterSelf(paragraph);
                insertPoint = paragraph;
            }
        }

        private IList<OpenXmlCompositeElement> ConvertContentToParagraphs(MainDocumentPart mainDocumentPart)
        {
            var converter = new HtmlConverter(mainDocumentPart) { RenderPreAsTable = false };
            return converter.Parse(_content
                .Replace("<pre><code>", "<p class=\"codesnippet\"><pre>")
                .Replace("</code></pre>", "</pre></p>")
                .Replace("<code>", "<span class=\"inlinecodesnippet\">")
                .Replace("</code>", "</span>")
                .Replace("<img", "<img width=\"550\""));
        }
    }
}
