using MarkdownSharp;
using SharpDox.Model.Documentation;
using SharpDox.Model.Repository;
using SharpDox.Plugins.Word.OpenXml;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;
using System.IO;

namespace SharpDox.Plugins.Word.Templaters
{
    internal class NamespaceTemplate : BaseTemplate
    {
        private readonly SDNamespace _sdNamespace;
        private readonly string _language;
        private readonly int _navigationLevel;

        public NamespaceTemplate(SDNamespace sdNamespace, string language, string outputPath, int navigationLevel) : base(outputPath, Templates.Namespace)
        {
            _sdNamespace = sdNamespace;
            _language = sdNamespace.Description.ContainsKey(language) ? language : "default";
            _navigationLevel = navigationLevel;
        }

        public override void CreateDocument()
        {
            ReplaceBookmarks();
            FillTable();
            CreateAndMergeTypeDocuments();
        }

        private void ReplaceBookmarks()
        {
            var description = string.Empty;
            _sdNamespace.Description.TryGetValue(_language, out description);

            var data = new List<FieldData>();
            data.Add(new FieldData("Fullname", new PlainText(_sdNamespace.Fullname)) { StyleName = string.Format("Heading {0}", _navigationLevel) });
            data.Add(new FieldData("Description", new RichText(new Markdown().Transform(description))));

            _templater.ReplaceBookmarks(data);
        }

        private void FillTable()
        {
            foreach (var sdType in _sdNamespace.Types)
            {
                SDDocumentation documentation;
                sdType.Documentation.TryGetValue(_language, out documentation);

                _templater.AddRowToTable(0, new[] {
                    sdType.Name,
                    documentation != null ? documentation.Summary.ToString() : string.Empty
                });
            }
        }

        private void CreateAndMergeTypeDocuments()
        {

        }
    }
}
