using SharpDox.Model;
using SharpDox.Model.Repository;
using SharpDox.Plugins.Word.OpenXml;

namespace SharpDox.Plugins.Word.Builder
{
    internal class ApiBuilder : BaseApiBuilder
    {
        private readonly NamespaceBuilder _namespaceBuilder;

        public ApiBuilder(WordTemplater wordTemplater, SDProject sdProject, WordStrings wordStrings, string docLanguage, string outputPath)
            : base(wordTemplater, sdProject, wordStrings, docLanguage) 
        {
            _namespaceBuilder = new NamespaceBuilder(wordTemplater, sdProject, wordStrings, docLanguage, outputPath);
        }

        public void CreateApiDoc(SDRepository sdRepository, int navigationLevel)
        {
            foreach (var sdNamespace in sdRepository.GetAllNamespaces())
            {
                _namespaceBuilder.InsertNamespace(sdNamespace, navigationLevel);
            }
        }
    }
}
