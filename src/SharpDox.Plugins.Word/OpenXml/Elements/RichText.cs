using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using NotesFor.HtmlToOpenXml;

namespace SharpDox.Plugins.Word.OpenXml.Elements
{
    internal class RichText : BaseElement
    {
        public RichText(string content) : base(content) { }

        public override void AppendTo(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            var converter = new HtmlConverter(mainDocumentPart) { RenderPreAsTable = false };
            var paragraphs = converter.Parse(_content.Replace("<pre>", "<p class=\"codesnippet\"><pre>").Replace("</pre>", "</pre></p>"));

            foreach (var paragraph in paragraphs)
            {
                openXmlNode.Append(paragraph);
            }
        }

        public override void InsertAfter(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            var converter = new HtmlConverter(mainDocumentPart) { RenderPreAsTable = false };
            var paragraphs = converter.Parse(_content.Replace("<pre>", "<p class=\"codesnippet\"><pre>").Replace("</pre>", "</pre></p>"));

            var insertPoint = openXmlNode;
            foreach (var paragraph in paragraphs)
            {
                insertPoint.InsertAfterSelf(paragraph);
                insertPoint = paragraph;
            }
        }
    }
}
