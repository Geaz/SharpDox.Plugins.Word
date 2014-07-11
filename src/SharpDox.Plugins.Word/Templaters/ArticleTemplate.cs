using MarkdownSharp;
using SharpDox.Model;
using SharpDox.Model.Documentation.Article;
using SharpDox.Plugins.Word.OpenXml;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;

namespace SharpDox.Plugins.Word.Templaters
{
    internal class ArticleTemplate : BaseTemplate
    {
        private readonly SDProject _sdProject;
        private readonly SDArticle _article;
        private readonly int _navigationLevel;

        public ArticleTemplate(SDProject sdProject, SDArticle article, string outputPath, int navigationLevel) : base(outputPath, Templates.Article)
        {
            _sdProject = sdProject;
            _article = article;
            _navigationLevel = navigationLevel;
        }

        public override void CreateDocument()
        {
            var data = new List<FieldData>();
            data.Add(new FieldData("Title", new PlainText(_article.Title)) { StyleName = string.Format("Heading {0}", _navigationLevel) });
            data.Add(new FieldData("Content", new RichText(new Markdown().Transform(_article.Content.Transform(new Helper(_sdProject).TransformLinkToken)))));
            _templater.ReplaceBookmarks(data);
        }
    }
}
