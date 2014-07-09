using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharpDox.Plugins.Word.OpenXml
{
    internal class FieldData
    {
        public FieldData(string fieldName, string text, bool isMarkdown = false)
        {
            FieldName = fieldName;
            Text = text;
            IsMarkDown = isMarkdown;
        }

        public string FieldName { get; set; }
        public string Text { get; set; }
        public string StyleName { get; set; }
        public bool IsMarkDown { get; set; }
    }
}
