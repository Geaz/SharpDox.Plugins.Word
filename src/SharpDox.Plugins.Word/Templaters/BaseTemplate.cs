using SharpDox.Plugins.Word.OpenXml;
using System.IO;

namespace SharpDox.Plugins.Word.Templaters
{
    internal abstract class BaseTemplate
    {
        protected readonly string _outputPath;
        protected readonly WordTemplater _templater;

        public BaseTemplate(string outputPath, string template)
        {
            _outputPath = outputPath;
            TemplatePath = Helper.EnsureCopy(
                                Path.Combine(Path.GetDirectoryName(GetType().Assembly.Location), Templates.Folder, template), 
                                Path.Combine(outputPath, "tmp"));
            _templater = new WordTemplater(TemplatePath);
        }

        public abstract void CreateDocument();

        public string TemplatePath { private set; get; }
    }
}
