using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SharpDox.Model;
using SharpDox.Model.Documentation.Article;
using SharpDox.Model.Repository;
using SharpDox.Plugins.Word.Templaters;
using SharpDox.Sdk.Exporter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharpDox.Plugins.Word
{
    public class Exporter : IExporter
    {
        public event Action<string> OnStepMessage;
        public event Action<int> OnStepProgress; 
        public event Action<string> OnRequirementsWarning;
                
        private SDProject _sdProject;
        private string _outputPath;
        private int _docCount;
        private int _docIndex;

        private string _currentOutputPath, _currentDocLanguage;
        private MainTemplate _mainTemplate;

        public void Export(SDProject sdProject, string outputPath)
        {
            _sdProject = sdProject;
            _outputPath = outputPath;
            _docCount = sdProject.DocumentationLanguages.Count;
            _docIndex = 0;

            foreach (var docLanguage in sdProject.DocumentationLanguages)
            {
                _currentOutputPath = Path.Combine(outputPath, docLanguage);
                _currentDocLanguage = docLanguage;

                _mainTemplate = new MainTemplate(sdProject, _currentDocLanguage, _currentOutputPath);
                _mainTemplate.CreateDocument();

                if(sdProject.Articles.Count != 0)
                {
                    List<SDArticle> articles;
                    sdProject.Articles.TryGetValue(_currentDocLanguage, out articles);
                    articles = articles ?? sdProject.Articles["default"];
                    CreateArticles(articles, false, 1);
                }
                else
                {
                    CreateApiDoc(_sdProject.Repositories.Single().Value, 1);
                }

                _mainTemplate.SaveToOutputFolder();
                Directory.Delete(Path.Combine(_currentOutputPath, "tmp"), true);
                _docIndex++;
            }
        }

        public bool CheckRequirements() { return true; }

        private void CreateArticles(IEnumerable<SDArticle> articles, bool nextMergeWithPageBreak, int navigationLevel)
        {
            foreach (var article in articles)
            {
                var placeholder = article as SDDocPlaceholder;
                if (placeholder != null)
                {
                    CreateApiDoc(_sdProject.Repositories[placeholder.SolutionFile], navigationLevel);
                    nextMergeWithPageBreak = true;
                }
                else if (article is SDArticlePlaceholder)
                {
                    var placeholderTemplate = new PlaceholderTemplate(article.Title, _currentOutputPath);
                    placeholderTemplate.CreateDocument();
                    _mainTemplate.MergeWith(placeholderTemplate.TemplatePath, nextMergeWithPageBreak);
                    nextMergeWithPageBreak = false;
                }
                else
                {
                    var articleTemplate = new ArticleTemplate(article, _currentOutputPath, navigationLevel);
                    articleTemplate.CreateDocument();
                    _mainTemplate.MergeWith(articleTemplate.TemplatePath, nextMergeWithPageBreak);
                    nextMergeWithPageBreak = true;
                }

                if(article.Children.Count > 0)
                {
                    CreateArticles(article.Children, nextMergeWithPageBreak, navigationLevel + 1);
                }
            }
        }

        private void CreateApiDoc(SDRepository sdRepository, int navigationLevel)
        {
            foreach (var sdNamespace in sdRepository.GetAllNamespaces())
            {
                var namespaceTemplate = new NamespaceTemplate(sdNamespace, _currentDocLanguage, _currentOutputPath, navigationLevel);
                namespaceTemplate.CreateDocument();
                _mainTemplate.MergeWith(namespaceTemplate.TemplatePath, true);
            }
        }

        private void ExecuteOnStepMessage(string message)
        {
            var handle = OnStepMessage;
            if (handle != null)
            {
                handle(message);
            }
        }

        private void ExecuteOnStepProgress(int progress)
        {
            var handle = OnStepProgress;
            if (handle != null)
            {
                handle(progress);
            }
        }

        public string ExporterName { get { return "Word"; } }
    }
}
