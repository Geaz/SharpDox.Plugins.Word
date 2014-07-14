using MarkdownSharp;
using SharpDox.Model;
using SharpDox.Model.Repository;
using SharpDox.Plugins.Word.OpenXml;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;

namespace SharpDox.Plugins.Word.Templaters
{
    internal class NamespaceTemplate : BaseTemplate
    {
        private readonly SDProject _sdProject;
        private readonly SDNamespace _sdNamespace;
        private readonly WordStrings _wordStrings;
        private readonly string _language;
        private readonly int _navigationLevel;

        public NamespaceTemplate(SDProject sdProject, SDNamespace sdNamespace, WordStrings wordStrings, string language, string outputPath, int navigationLevel) 
            : base(outputPath, Templates.Namespace)
        {
            _sdProject = sdProject;
            _sdNamespace = sdNamespace;
            _wordStrings = wordStrings;
            _language = language;
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
            var description = _sdNamespace.Descriptions.GetElementOrDefault(_language);
            var data = new List<FieldData>();
            data.Add(new FieldData("Fullname", new PlainText(_sdNamespace.Fullname)) { StyleName = string.Format("Heading {0}", _navigationLevel) });
            data.Add(new FieldData("Description", new RichText(description != null ? new Markdown().Transform(description.Transform(new Helper(_sdProject).TransformLinkToken)) : string.Empty)));
            data.Add(new FieldData("Overview_Text", new PlainText(_wordStrings.Overview)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("Name_Text", new PlainText(_wordStrings.Name)));
            data.Add(new FieldData("Description_Text", new PlainText(_wordStrings.Description)));

            _templater.ReplaceBookmarks(data);
        }

        private void FillTable()
        {
            foreach (var sdType in _sdNamespace.Types)
            {
                var documentation = sdType.Documentations.GetElementOrDefault(_language);
                _templater.AddRowToTable(0, new List<BaseElement> {
                    new Image(Icons.GetIconPath("class", sdType.Accessibility)),
                    new PlainText(sdType.Name),
                    new PlainText(documentation != null ? documentation.Summary.ToString() : string.Empty)
                });
            }
        }

        private void CreateAndMergeTypeDocuments()
        {
            foreach(var sdType in _sdNamespace.Types)
            {

            }
        }
    }
}
