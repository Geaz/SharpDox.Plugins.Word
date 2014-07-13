using SharpDox.Sdk.Local;

namespace SharpDox.Plugins.Word
{
    public class WordStrings : ILocalStrings
    {
        private string _title = "Title";
        private string _name = "Name";
        private string _description = "Description";
        private string _types = "Types";
        private string _disclaimer = "Created by sharpDox (http://sharpdox.de)";
        private string _tocHeader = "Table of contents";
        private string _tocBodyPlaceholder = "Please refresh the table.";

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

        public string Disclaimer
        {
            get { return _disclaimer; }
            set { _disclaimer = value; }
        }

        public string TocHeader
        {
            get { return _tocHeader; }
            set { _tocHeader = value; }
        }

        public string TocBodyPlaceholder
        {
            get { return _tocBodyPlaceholder; }
            set { _tocBodyPlaceholder = value; }
        }

        public string DisplayName { get { return "WordExporter"; } }
    }
}
