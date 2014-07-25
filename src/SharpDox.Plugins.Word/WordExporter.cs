using SharpDox.Model;
using SharpDox.Plugins.Word.Builder;
using SharpDox.Sdk.Exporter;
using SharpDox.Sdk.Local;
using System;
using System.IO;

namespace SharpDox.Plugins.Word
{
    public class WordExporter : IExporter
    {
        public event Action<string> OnStepMessage;
        public event Action<int> OnStepProgress; 
        public event Action<string> OnRequirementsWarning;
                
        private SDProject _sdProject;
        private string _outputPath, _currentDocLanguage;
        private double _docCount, _docIndex;

        private readonly ILocalController _localController;
        private readonly WordStrings _wordStrings;

        public WordExporter(ILocalController localController)
        {
            _localController = localController;
            _wordStrings = _localController.GetLocalStrings<WordStrings>();
        }

        public void Export(SDProject sdProject, string outputPath)
        {
            _sdProject = sdProject;
            _outputPath = outputPath;
            _docCount = sdProject.DocumentationLanguages.Count;
            _docIndex = 0;

            foreach (var docLanguage in sdProject.DocumentationLanguages)
            {
                var currentOutputPath = Path.Combine(outputPath, docLanguage);
                _currentDocLanguage = docLanguage;

                ExecuteOnStepMessage(_wordStrings.LoadingTemplate);
                ExecuteOnStepProgress(10);
                var docBuilder = new DocBuilder(_sdProject, _localController.GetLocalStringsOrDefault<WordStrings>(_currentDocLanguage), _currentDocLanguage, currentOutputPath);

                ExecuteOnStepProgress(20);
                ExecuteOnStepMessage(_wordStrings.CreatingDocument);
                docBuilder.BuildDocument();

                ExecuteOnStepProgress(80);
                ExecuteOnStepMessage(_wordStrings.SavingDocument);
                docBuilder.SaveToOutputFolder();

                ExecuteOnStepProgress(90);
                ExecuteOnStepMessage(_wordStrings.DeleteTmp);
                Directory.Delete(Path.Combine(_outputPath, _currentDocLanguage, "tmp"), true);

                _docIndex++;
            }
        }

        public bool CheckRequirements() { return true; }

        private void ExecuteOnStepMessage(string message)
        {
            var handle = OnStepMessage;
            if (handle != null)
            {
                handle(string.Format("({0}) - {1}", _currentDocLanguage, message));
            }
        }

        private void ExecuteOnStepProgress(int progress)
        {
            var handle = OnStepProgress;
            if (handle != null)
            {
                handle((int)((progress / _docCount) + (100 / _docCount * _docIndex)));
            }
        }

        public string ExporterName { get { return "Word"; } }
    }
}
