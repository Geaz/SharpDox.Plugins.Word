using SharpDox.Sdk.Local;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharpDox.Plugins.Word
{
    public class WordStrings : ILocalStrings
    {
        private string _title = "Title";
        private string _name = "Name";
        private string _description = "Description";
        private string _types = "Types";

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

        public string DisplayName { get { return "WordExporter"; } }
    }
}
