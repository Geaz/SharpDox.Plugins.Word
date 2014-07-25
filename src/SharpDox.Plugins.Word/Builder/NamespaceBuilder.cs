using System.Text;
using MarkdownSharp;
using SharpDox.Model;
using SharpDox.Model.Repository;
using SharpDox.Plugins.Word.OpenXml;
using System.Collections.Generic;
using System.Linq;

namespace SharpDox.Plugins.Word.Builder
{
    internal class NamespaceBuilder : BaseApiBuilder
    {
        private readonly TypeBuilder _typeBuilder;

        public NamespaceBuilder(WordTemplater wordTemplater, SDProject sdProject, WordStrings wordStrings, string docLanguage, string outputPath)
            : base(wordTemplater, sdProject, wordStrings, docLanguage) 
        {
            _typeBuilder = new TypeBuilder(wordTemplater, sdProject, wordStrings, docLanguage, outputPath);
        }

        public void InsertNamespace(SDNamespace sdNamespace, int navigationLevel)
        {
            _wordTemplater.AppendHeader(sdNamespace.Fullname, navigationLevel);
            InsertNamespaceDescription(sdNamespace);
            InsertNamespaceOverview(sdNamespace, navigationLevel + 1);
            InsertUseUsedBlock(sdNamespace, navigationLevel + 1);
            _wordTemplater.AppendPageBreak();

            foreach (var sdType in sdNamespace.Types)
            {
                _typeBuilder.InsertType(sdType, navigationLevel + 1);
            }
        }

        private void InsertNamespaceDescription(SDNamespace sdNamespace)
        {
            var description = sdNamespace.Descriptions.GetElementOrDefault(_docLanguage);
            if (description != null)
            {
                _wordTemplater.AppendMarkdown(new Markdown().Transform(description.Transform(new Helper(_sdProject).TransformLinkToken)));
            }
        }

        private void InsertNamespaceOverview(SDNamespace sdNamespace, int navigationLevel)
        {
            _wordTemplater.AppendHeader(_wordStrings.Overview, navigationLevel);

            var headers = new List<string> {
                    string.Empty,
                    _wordStrings.Name,
                    _wordStrings.Description
                };

            var rows = new List<List<string>>();
            foreach (var sdType in sdNamespace.Types)
            {
                var documentation = sdType.Documentations.GetElementOrDefault(_docLanguage);
                rows.Add(new List<string> {
                        string.Format("<img width=\"16\" src=\"{0}\"/>", Icons.GetIconPath("class", sdType.Accessibility)),
                        sdType.Name,
                        documentation != null ? documentation.Summary.ToString() : string.Empty
                    });
            }

            _wordTemplater.AppendTable(headers, rows);
        }

        private void InsertUseUsedBlock(SDNamespace sdNamespace, int navigationLevel)
        {
            if (sdNamespace.Uses.Count > 0)
            {
                _wordTemplater.AppendHeader(_wordStrings.Uses, navigationLevel);
                sdNamespace.Uses.Select(u => u.Fullname).ToList().ForEach(u => _wordTemplater.AppendParagraph(u, "CenteredNoMargin"));
            }

            if (sdNamespace.UsedBy.Count > 0)
            {
                _wordTemplater.AppendHeader(_wordStrings.UsedBy, navigationLevel);
                sdNamespace.UsedBy.Select(u => u.Fullname).ToList().ForEach(u => _wordTemplater.AppendParagraph(u, "CenteredNoMargin"));
            }
        }
    }
}
