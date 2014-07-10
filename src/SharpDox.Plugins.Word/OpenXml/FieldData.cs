using SharpDox.Plugins.Word.OpenXml.Elements;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
