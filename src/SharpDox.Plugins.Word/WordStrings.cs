using SharpDox.Sdk.Local;

namespace SharpDox.Plugins.Word
{
    public class WordStrings : ILocalStrings
    {
        private string _loadingTemplate = "Loading template";
        private string _creatingDocument = "Creating document";
        private string _savingDocument = "Saving Document";
        private string _deleteTmp = "Deleting temporary folder";

        private string _title = "Title";
        private string _name = "Name";
        private string _description = "Description";
        private string _types = "Types";
        private string _overview = "Overview";
        private string _disclaimer = "Created by sharpDox (http://sharpdox.de)";
        private string _remarks = "Remarks";
        private string _returns = "Returns";
        private string _example = "Example";
        private string _exceptions = "Exceptions";
        private string _typeParams = "Type Parameters";
        private string _params = "Paramters";
        private string _seeAlso = "See also";
        private string _uses = "Uses";
        private string _usedBy = "Used By";
        private string _fields = "Fields";
        private string _events = "Events";
        private string _methods = "Methods";
        private string _properties = "Properties";
        private string _tocPlaceholder = "Please insert a table of contents.";

        public string LoadingTemplate
        {
            get { return _loadingTemplate; }
            set { _loadingTemplate = value; }
        }

        public string CreatingDocument
        {
            get { return _creatingDocument; }
            set { _creatingDocument = value; }
        }

        public string SavingDocument
        {
            get { return _savingDocument; }
            set { _savingDocument = value; }
        }

        public string DeleteTmp
        {
            get { return _deleteTmp; }
            set { _deleteTmp = value; }
        }

        public string Title
        {
            get { return _title; }
            set { _title = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public string Description
        {
            get { return _description; }
            set { _description = value; }
        }

        public string Types
        {
            get { return _types; }
            set { _types = value; }
        }

        public string Overview
        {
            get { return _overview; }
            set { _overview = value; }
        }

        public string Disclaimer
        {
            get { return _disclaimer; }
            set { _disclaimer = value; }
        }

        public string TocPlaceholder
        {
            get { return _tocPlaceholder; }
            set { _tocPlaceholder = value; }
        }

        public string Remarks
        {
            get { return _remarks; }
            set { _remarks = value; }
        }

        public string Returns
        {
            get { return _returns; }
            set { _returns = value; }
        }

        public string Example
        {
            get { return _example; }
            set { _example = value; }
        }

        public string Exceptions
        {
            get { return _exceptions; }
            set { _exceptions = value; }
        }

        public string Params
        {
            get { return _params; }
            set { _params = value; }
        }

        public string TypeParams
        {
            get { return _typeParams; }
            set { _typeParams = value; }
        }

        public string SeeAlso
        {
            get { return _seeAlso; }
            set { _seeAlso = value; }
        }

        public string Uses
        {
            get { return _uses; }
            set { _uses = value; }
        }

        public string UsedBy
        {
            get { return _usedBy; }
            set { _usedBy = value; }
        }

        public string Fields
        {
            get { return _fields; }
            set { _fields = value; }
        }

        public string Events
        {
            get { return _events; }
            set { _events = value; }
        }

        public string Methods
        {
            get { return _methods; }
            set { _methods = value; }
        }

        public string Properties
        {
            get { return _properties; }
            set { _properties = value; }
        }

        public string DisplayName { get { return "WordExporter"; } }
    }
}
