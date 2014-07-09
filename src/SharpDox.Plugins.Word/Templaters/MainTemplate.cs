using SharpDox.Model;
using SharpDox.Plugins.Word.OpenXml;
using System;
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
            data.Add(new FieldData("Title", _sdProject.ProjectName));
            data.Add(new FieldData("Version", _sdProject.VersionNumber));
            data.Add(new FieldData("Author", _sdProject.Author));
            data.Add(new FieldData("Date", DateTime.Now.ToShortDateString()));
            data.Add(new FieldData("AuthorUrl", _sdProject.AuthorUrl));
            data.Add(new FieldData("ProjectUrl", _sdProject.ProjectUrl));
            _templater.ReplaceBookmarks(data);
        }

        public void MergeWith(string documentPath, bool pageBreak)
        {
            if (pageBreak)
            {
                WordMerger.MergeDocumentsWithPagebreak(TemplatePath, documentPath);
            }
            else
            {
                WordMerger.MergeDocuments(TemplatePath, documentPath);
            }
        }

        public void SaveToOutputFolder()
        {
            File.Copy(TemplatePath, Path.Combine(_outputPath, string.Format("{0}-{1}.docx", _sdProject.ProjectName, _language)), true);
        }
    }
}
