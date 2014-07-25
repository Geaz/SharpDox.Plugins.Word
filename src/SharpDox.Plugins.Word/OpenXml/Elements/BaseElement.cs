using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace SharpDox.Plugins.Word.OpenXml.Elements
{
    internal abstract class BaseElement
    {
        protected readonly string _content;

        public BaseElement(string content)
        {
            _content = content;
        }

        public abstract void AppendTo(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart);
        public abstract void InsertAfter(OpenXmlElement openXmlNode, MainDocumentPart mainDocumentPart);

        protected string GetStyleIdbyName(MainDocumentPart mainDocumentPart, string styleName)
        {
            var style = (Style)mainDocumentPart.StyleDefinitionsPart.Styles.FirstOrDefault(s => s is Style && ((Style)s).StyleName.Val.Value.ToLower() == styleName.ToLower());
            return style != null ? style.StyleId.Value : string.Empty;
        }
    }
}
