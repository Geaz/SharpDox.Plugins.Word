using System.Text;
using SharpDox.Model;
using SharpDox.Model.Documentation;
using SharpDox.Model.Documentation.Token;
using SharpDox.Model.Repository;
using SharpDox.Model.Repository.Members;
using SharpDox.Plugins.Word.OpenXml;
using System.Collections.Generic;
using System.Linq;
using SharpDox.UML;
using System.IO;

namespace SharpDox.Plugins.Word.Builder
{
    internal class TypeBuilder : BaseApiBuilder
    {
        private readonly string _outputPath;

        public TypeBuilder(WordTemplater wordTemplater, SDProject sdProject, WordStrings wordStrings, string docLanguage, string outputPath)
            : base(wordTemplater, sdProject, wordStrings, docLanguage) 
        {
            _outputPath = outputPath;
        }

        public void InsertType(SDType sdType, int navigationLevel)
        {
            var documentation = sdType.Documentations.GetElementOrDefault(_docLanguage);
            _wordTemplater.AppendHeader(sdType.Name, navigationLevel);
            _wordTemplater.AppendCodeBlock(sdType.Syntax);
            InsertDocumentation(null, documentation, navigationLevel + 1);
            InsertClassDiagram(sdType);
            InsertUseUsedBlock(sdType, navigationLevel + 1);

            InsertMemberBlock(_wordStrings.Fields, "field", sdType.Fields.Cast<SDMember>(), navigationLevel + 1);
            InsertMemberBlock(_wordStrings.Events, "event", sdType.Events.Cast<SDMember>(), navigationLevel + 1);
            InsertMemberBlock(_wordStrings.Methods, "method", sdType.Methods.Cast<SDMember>(), navigationLevel + 1);
            InsertMemberBlock(_wordStrings.Properties, "properties", sdType.Properties.Cast<SDMember>(), navigationLevel + 1);
            _wordTemplater.AppendPageBreak();
        }

        private void InsertClassDiagram(SDType sdType)
        {
            if(!sdType.IsClassDiagramEmpty())
            {
                var tmpImagePath = Path.Combine(_outputPath, "tmp", sdType.Guid + ".png");
                sdType.GetClassDiagram().ToPng(tmpImagePath);
                _wordTemplater.AppendImage(tmpImagePath, "Diagram");
            }
        }

        private void InsertUseUsedBlock(SDType sdType, int navigationLevel)
        {
            if (sdType.Uses.Count > 0)
            {
                _wordTemplater.AppendHeader(_wordStrings.Uses, navigationLevel);
                sdType.Uses.Select(u => u.Fullname).ToList().ForEach(u => _wordTemplater.AppendParagraph(u, "CenteredNoMargin"));
            }

            if (sdType.UsedBy.Count > 0)
            {
                _wordTemplater.AppendHeader(_wordStrings.UsedBy, navigationLevel);
                sdType.UsedBy.Select(u => u.Fullname).ToList().ForEach(u => _wordTemplater.AppendParagraph(u, "CenteredNoMargin"));
            }
        }

        private void InsertMemberBlock(string title, string memberType, IEnumerable<SDMember> members, int navigationLevel)
        {
            if (members.Count() > 0)
            {
                _wordTemplater.AppendHeader(title, navigationLevel);
                foreach (var member in members)
                {
                    InsertMember(memberType, member, navigationLevel + 1);
                }
            }
        }

        private void InsertMember(string memberType, SDMember sdMember, int navigationLevel)
        {
            var documentation = sdMember.Documentations.GetElementOrDefault(_docLanguage);
            _wordTemplater.AppendRichText(string.Format("<p class=\"MemberHeader\"><img width=\"16\" src=\"{0}\"/> {1}</p>", Icons.GetIconPath(memberType, sdMember.Accessibility), sdMember.Name));
            _wordTemplater.AppendCodeBlock(sdMember.Syntax);
            InsertDocumentation(sdMember, documentation, navigationLevel + 1);
        }

        private void InsertDocumentation(SDMember sdMember, SDDocumentation documentation, int navigationLevel)
        {
            if (documentation != null)
            {
                if (documentation.Summary != null) _wordTemplater.AppendMarkdown(documentation.Summary.ToMarkdown(), "MemberBody");
                InsertDocumentationPart(_wordStrings.Remarks, documentation.Remarks);
                InsertDocumentationPart(_wordStrings.Returns, documentation.Returns);
                InsertParamDocumentationPart(_wordStrings.Params, documentation.Params, sdMember);
                InsertParamDocumentationPart(_wordStrings.TypeParams, documentation.TypeParams, sdMember);
                InsertParamDocumentationPart(_wordStrings.Exceptions, documentation.Exceptions, sdMember);
                InsertDocumentationPart(_wordStrings.Example, documentation.Example);
            }
        }

        private void InsertDocumentationPart(string title, SDTokenList documentation)
        {
            if (documentation != null && documentation.Count > 0)
            {
                _wordTemplater.AppendParagraph(title, "MemberSubHeader");
                _wordTemplater.AppendMarkdown(documentation.ToMarkdown(), "MemberBody");
            }
        }

        private void InsertParamDocumentationPart(string title, Dictionary<string, SDTokenList> parameters, SDMember sdMember)
        {
            if (parameters != null && parameters.Count > 0)
            {
                _wordTemplater.AppendParagraph(title, "MemberSubHeader");
                foreach (var parameter in parameters)
                {
                    _wordTemplater.AppendParagraph(parameter.Key, "Parameter");
                    if (sdMember is SDMethod)
                    {
                        var sdParam = ((SDMethod)sdMember).Parameters.SingleOrDefault(s => s.Name == parameter.Key);
                        if (sdParam != null)
                        {
                            _wordTemplater.AppendParagraph(sdParam.ParamType.Name, "ParameterType");
                        }
                    }
                    _wordTemplater.AppendMarkdown(parameter.Value.ToMarkdown(), "ParameterBody");
                }
            }
        }
    }
}
