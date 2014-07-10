using SharpDox.Plugins.Word.OpenXml;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;

namespace SharpDox.Plugins.Word.Templaters
{
    internal class PlaceholderTemplate : BaseTemplate
    {
        private readonly string _title;
        private readonly int _navigationLevel;

        public PlaceholderTemplate(string title, string outputPath, int navigationLevel) : base(outputPath, Templates.Placeholder)
        {
            _title = title;
            _navigationLevel = navigationLevel;
        }

        public override void CreateDocument()
        {
            var data = new List<FieldData>();
            data.Add(new FieldData("Title", new PlainText(_title)) { StyleName = string.Format("Heading {0}", _navigationLevel) });
            _templater.ReplaceBookmarks(data);
        }
    }
}
