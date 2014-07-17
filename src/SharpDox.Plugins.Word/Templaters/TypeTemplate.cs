using MarkdownSharp;
using SharpDox.Model.Repository;
using SharpDox.Model.Repository.Members;
using SharpDox.Plugins.Word.OpenXml;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;

namespace SharpDox.Plugins.Word.Templaters
{
    internal class TypeTemplate : BaseTemplate
    {
        private readonly SDType _sdType;
        private readonly WordStrings _wordStrings;
        private readonly string _language;
        private readonly int _navigationLevel;

        public TypeTemplate(SDType sdType, WordStrings wordStrings, string language, string outputPath, int navigationLevel) 
            : base(outputPath, Templates.Type)
        {
            _sdType = sdType;
            _wordStrings = wordStrings;
            _language = language;
            _navigationLevel = navigationLevel;
        }

        public override void CreateDocument()
        {
            ReplaceBookmarks();
            FillTable(_sdType.Fields, "field", 0);
            FillTable(_sdType.Events, "event", 1);
            FillTable(_sdType.Methods, "method", 2);
            FillTable(_sdType.Properties, "properties", 3);
            CreateAndMergeMemberDocuments();
        }

        private void ReplaceBookmarks()
        {
            var documentation = _sdType.Documentations.GetElementOrDefault(_language);
            var data = new List<FieldData>();
            data.Add(new FieldData("Typename", new PlainText(_sdType.Name)) { StyleName = string.Format("Heading {0}", _navigationLevel) });
            data.Add(new FieldData("Syntax", new PlainText(_sdType.Syntax)));
            data.Add(new FieldData("Summary", new RichText(documentation != null ? new Markdown().Transform(documentation.Summary.ToMarkdown()) : string.Empty)));
            data.Add(new FieldData("Remarks_Text", new PlainText(_wordStrings.Remarks)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("Remarks", new RichText(documentation != null ? new Markdown().Transform(documentation.Remarks.ToMarkdown()) : string.Empty)));
            data.Add(new FieldData("Example_Text", new PlainText(_wordStrings.Example)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("Exceptions_Text", new PlainText(_wordStrings.Exceptions)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("TypeParams_Text", new PlainText(_wordStrings.TypeParams)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("SeeAlso_Text", new PlainText(_wordStrings.SeeAlso)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("Uses_Text", new PlainText(_wordStrings.Uses)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("UsedBy_Text", new PlainText(_wordStrings.UsedBy)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("Fields_Text", new PlainText(_wordStrings.Fields)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("Events_Text", new PlainText(_wordStrings.Events)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("Methods_Text", new PlainText(_wordStrings.Methods)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("Properties_Text", new PlainText(_wordStrings.Properties)) { StyleName = string.Format("Heading {0}", _navigationLevel + 1) });
            data.Add(new FieldData("Name_Text", new PlainText(_wordStrings.Name)));
            data.Add(new FieldData("Description_Text", new PlainText(_wordStrings.Description)));

            _templater.ReplaceBookmarks(data);
        }

        private void FillTable(IEnumerable<SDMember> members, string memberType, int tableIndex)
        {
            foreach (var member in members)
            {
                var documentation = member.Documentations.GetElementOrDefault(_language);
                _templater.AddRowToTable(tableIndex, new List<BaseElement> {
                    new Image(Icons.GetIconPath(memberType, member.Accessibility)),
                    new PlainText(member.Name),
                    new PlainText(documentation != null ? documentation.Summary.ToString() : string.Empty)
                });
            }
        }

        private void CreateAndMergeMemberDocuments()
        {
            
        }
    }
}
