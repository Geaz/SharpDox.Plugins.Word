using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace SharpDox.Plugins.Word.OpenXml.Elements
{
    internal abstract class BaseElement
    {
        protected readonly string _content;

        public BaseElement(string content)
        {
            _content = content;
        }

        public abstract void InsertAfter(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart);
    }
}
