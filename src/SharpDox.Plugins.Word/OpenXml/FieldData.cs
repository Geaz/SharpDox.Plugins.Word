using SharpDox.Plugins.Word.OpenXml.Elements;

namespace SharpDox.Plugins.Word.OpenXml
{
    internal class FieldData
    {
        public FieldData(string fieldName, BaseElement element)
        {
            FieldName = fieldName;
            Element = element;
        }

        public string FieldName { get; set; }
        public BaseElement Element { get; set; }
        public string StyleName { get; set; }
    }
}
