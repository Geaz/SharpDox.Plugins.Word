using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace SharpDox.Plugins.Word.OpenXml.Elements
{
    internal class PlainText : BaseElement
    {
        public PlainText(string content) : base(content) { }

        public override void InsertAfter(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart)
        {
            openXmlNode.InsertAfterSelf(new Run(new Text(_content)));
        }
    }
}
