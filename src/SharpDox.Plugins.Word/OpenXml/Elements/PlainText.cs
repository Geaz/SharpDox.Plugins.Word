using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SharpDox.Plugins.Word.OpenXml.Elements
{
    internal class PlainText : BaseElement
    {
        private readonly bool _newParagraph;

        public PlainText(string content, bool newParagraph = false) : base(content) 
        {
            _newParagraph = newParagraph;
        }

        public override void AppendTo(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            var text = _newParagraph 
                ? (OpenXmlElement)new Paragraph(new Run(new Text(_content)))
                : (OpenXmlElement)new Run(new Text(_content));

            openXmlNode.Append(text);
        }

        public override void InsertAfter(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            openXmlNode.InsertAfterSelf(new Run(new Text(_content)));
        }
    }
}
