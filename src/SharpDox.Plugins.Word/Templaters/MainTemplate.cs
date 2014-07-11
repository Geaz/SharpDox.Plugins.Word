using SharpDox.Model;
using SharpDox.Plugins.Word.OpenXml;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;
using System.IO;

namespace SharpDox.Plugins.Word.Templaters
{
    internal class MainTemplate : BaseTemplate
    {
        private readonly SDProject _sdProject;
        private readonly string _language;

        public MainTemplate(SDProject sdProject, string language, string outputPath) : base(outputPath, Templates.Main)
        {
            _sdProject = sdProject;
            _language = sdProject.Description.ContainsKey(language) ? language : "default";
        }

        public override void CreateDocument()
        {
            // not used at the moment - maybe a own site for the description in the future
            //var description = string.Empty;
            //_sdProject.Description.TryGetValue(_language, out description);

            var data = new List<FieldData>();
            data.Add(new FieldData("Title", string.IsNullOrEmpty(_sdProject.LogoPath) ? (BaseElement)new PlainText(_sdProject.ProjectName) : (BaseElement)new Image(_sdProject.LogoPath)));
            data.Add(new FieldData("Version", new PlainText(_sdProject.VersionNumber)));
            data.Add(new FieldData("Author", new PlainText(_sdProject.Author)));
            data.Add(new FieldData("AuthorUrl", new PlainText(_sdProject.AuthorUrl)));
            data.Add(new FieldData("ProjectUrl", new PlainText(_sdProject.ProjectUrl)));
            data.Add(new FieldData("Disclaimer", new PlainText("This document was created by sharpDox")));
            data.Add(new FieldData("Header", new PlainText(string.Format("{0} {1}", _sdProject.ProjectName, _sdProject.VersionNumber))));
            _templater.ReplaceBookmarks(data);
        }

        public void MergeWith(string documentPath, bool pageBreak)
        {
            if (pageBreak)
            {
                WordMerger.MergeDocumentsWithPagebreak(documentPath, TemplatePath);
            }
            else
            {
                WordMerger.MergeDocuments(documentPath, TemplatePath);
            }
        }

        public void SaveToOutputFolder()
        {
            File.Copy(TemplatePath, Path.Combine(_outputPath, string.Format("{0}-{1}.docx", _sdProject.ProjectName, _language)), true);
        }
    }
}
