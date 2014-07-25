using SharpDox.Model;
using SharpDox.Model.Documentation.Article;
using SharpDox.Plugins.Word.OpenXml;
using System.Collections.Generic;

namespace SharpDox.Plugins.Word.Builder
{
    internal class ArticleBuilder
    {
        private readonly WordTemplater _wordTemplater;
        private readonly SDProject _sdProject;
        private readonly ApiBuilder _apiBuilder;

        public ArticleBuilder(WordTemplater wordTemplater, SDProject sdProject, ApiBuilder apiBuilder)
        {
            _wordTemplater = wordTemplater;
            _sdProject = sdProject;
            _apiBuilder = apiBuilder;
        }

        public void CreateArticles(IEnumerable<SDArticle> articles, int navigationLevel)
        {
            foreach (var article in articles)
            {
                var placeholder = article as SDDocPlaceholder;
                if (placeholder != null)
                {
                    _wordTemplater.AppendHeader(article.Title, navigationLevel);
                    _apiBuilder.CreateApiDoc(_sdProject.Repositories[placeholder.SolutionFile], navigationLevel + 1);
                }
                else if (article is SDArticlePlaceholder)
                {
                    _wordTemplater.AppendHeader(article.Title, navigationLevel);
                }
                else
                {
                    _wordTemplater.AppendHeader(article.Title, navigationLevel);
                    _wordTemplater.AppendMarkdown(article.Content.Transform(new Helper(_sdProject).TransformLinkToken));
                    _wordTemplater.AppendPageBreak();
                }

                if (article.Children.Count > 0)
                {
                    CreateArticles(article.Children, navigationLevel + 1);
                }
            }
        }
    }
}
