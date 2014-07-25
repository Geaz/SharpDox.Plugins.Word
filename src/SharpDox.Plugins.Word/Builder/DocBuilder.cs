using SharpDox.Model;
using SharpDox.Model.Documentation.Article;
using SharpDox.Plugins.Word.OpenXml;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharpDox.Plugins.Word.Builder
{
    internal class DocBuilder
    {
        private readonly SDProject _sdProject;
        private readonly WordStrings _wordStrings;
        private readonly string _docLanguage, _outputPath, _templatePath;
        private readonly WordTemplater _wordTemplater;

        private readonly ApiBuilder _apiBuilder;
        private readonly ArticleBuilder _articleBuilder;

        public DocBuilder(SDProject sdProject, WordStrings wordStrings, string docLanguage, string outputPath)
        {
            _sdProject = sdProject;
            _wordStrings = wordStrings;
            _docLanguage = docLanguage;
            _outputPath = outputPath;

            _templatePath = Helper.EnsureCopy(
                                Path.Combine(Path.GetDirectoryName(GetType().Assembly.Location), "templates", "sharpDox.docx"),
                                Path.Combine(outputPath, "tmp"));
            _wordTemplater = new WordTemplater(_templatePath);

            _apiBuilder = new ApiBuilder(_wordTemplater, _sdProject, _wordStrings, _docLanguage, outputPath);
            _articleBuilder = new ArticleBuilder(_wordTemplater, _sdProject, _apiBuilder);
        }

        public void BuildDocument()
        {
            InitDocument();

            if (_sdProject.Articles.Count != 0)
            {
                List<SDArticle> articles;
                _sdProject.Articles.TryGetValue(_docLanguage, out articles);
                articles = articles ?? _sdProject.Articles["default"];
                _articleBuilder.CreateArticles(articles, 1);
            }
            else
            {
                _apiBuilder.CreateApiDoc(_sdProject.Repositories.Single().Value, 1);
            }

            _wordTemplater.Close();
        }

        public void SaveToOutputFolder()
        {
            File.Copy(_templatePath, Path.Combine(_outputPath, string.Format("{0}-{1}.docx", _sdProject.ProjectName, _docLanguage)), true);
        }

        private void InitDocument()
        {
            var data = new List<BookmarkData>();
            data.Add(new BookmarkData("Title", string.IsNullOrEmpty(_sdProject.LogoPath) ? (BaseElement)new PlainText(_sdProject.ProjectName) : (BaseElement)new Image(_sdProject.LogoPath)));
            data.Add(new BookmarkData("Version", new PlainText(_sdProject.VersionNumber)));
            data.Add(new BookmarkData("Author", new PlainText(_sdProject.Author)));
            data.Add(new BookmarkData("AuthorUrl", new PlainText(_sdProject.AuthorUrl)));
            data.Add(new BookmarkData("ProjectUrl", new PlainText(_sdProject.ProjectUrl)));
            data.Add(new BookmarkData("TocPlaceholder", new PlainText(_wordStrings.TocPlaceholder)));
            data.Add(new BookmarkData("Disclaimer", new PlainText(_wordStrings.Disclaimer)) { StyleName = "HeaderFooter" });
            data.Add(new BookmarkData("Header", new PlainText(string.Format("{0} {1}", _sdProject.ProjectName, _sdProject.VersionNumber))) { StyleName = "HeaderFooter" });
            _wordTemplater.ReplaceBookmarks(data);
        }
    }
}
