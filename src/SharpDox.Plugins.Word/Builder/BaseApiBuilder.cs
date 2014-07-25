using SharpDox.Model;
using SharpDox.Plugins.Word.OpenXml;

namespace SharpDox.Plugins.Word.Builder
{
    internal abstract class BaseApiBuilder
    {
        protected readonly WordTemplater _wordTemplater;
        protected readonly SDProject _sdProject;
        protected readonly WordStrings _wordStrings;
        protected readonly string _docLanguage;

        public BaseApiBuilder(WordTemplater wordTemplater, SDProject sdProject, WordStrings wordStrings, string docLanguage)
        {
            _wordTemplater = wordTemplater;
            _sdProject = sdProject;
            _wordStrings = wordStrings;
            _docLanguage = docLanguage;
        }
    }
}
