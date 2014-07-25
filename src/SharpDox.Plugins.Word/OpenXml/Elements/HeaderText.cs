using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace SharpDox.Plugins.Word.OpenXml.Elements
{
    internal class HeaderText : BaseElement
    {
        private readonly int _navigationLevel;

        public HeaderText(string content, int navigationLevel) : base(content) 
        {
            _navigationLevel = navigationLevel;
        }

        public override void AppendTo(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            openXmlNode.Append(GetHeaderTextElement(mainDocumentPart));
        }

        public override void InsertAfter(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            openXmlNode.InsertAfterSelf(GetHeaderTextElement(mainDocumentPart));
        }

        private OpenXmlElement GetHeaderTextElement(MainDocumentPart mainDocumentPart)
        {
            var paragraph = new Paragraph(new Run(new Text(_content)));
            var styleId = GetStyleIdbyName(mainDocumentPart, string.Format("Heading {0}", _navigationLevel));
            if (!string.IsNullOrEmpty(styleId))
            {
                paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = styleId });
            }

            return paragraph;
        }
    }
}
