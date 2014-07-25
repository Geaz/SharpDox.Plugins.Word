using SharpDox.Plugins.Word.OpenXml.Elements;

namespace SharpDox.Plugins.Word.OpenXml
{
    internal class BookmarkData
    {
        public BookmarkData(string bookmarkName, BaseElement element)
        {
            BookmarkName = bookmarkName;
            Element = element;
        }

        public string BookmarkName { get; set; }
        public BaseElement Element { get; set; }
        public string StyleName { get; set; }
    }
}
