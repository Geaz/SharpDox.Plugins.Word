using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace SharpDox.Plugins.Word.OpenXml
{
    internal class BookmarkReplacer
    {
        private readonly WordprocessingDocument _document;

        public BookmarkReplacer(WordprocessingDocument document)
        {
            _document = document;
        }

        public void ReplaceBookmarks(IEnumerable<BookmarkData> bookmarkData)
        {
            foreach (var data in bookmarkData)
            {
                ReplaceBookmark(data);
            }
        }

        public void ReplaceBookmark(BookmarkData bookmarkData)
        {
            var bookmarks = GetAllBookmarks(_document);
            if (bookmarks.ContainsKey(bookmarkData.BookmarkName))
            {
                foreach (var bookmark in bookmarks[bookmarkData.BookmarkName])
                {
                    DeleteBookmarkContent(bookmark);

                    if (bookmarkData.Element is RichText)
                    {
                        bookmarkData.Element.InsertAfter(bookmark.Parent, _document.MainDocumentPart);
                        // Richtext already includes style information (converted from html), no need to add it

                        // Bookmarks are included in Paragraphs
                        // But paragraphs included in paragraphs are not allowed
                        // Thats why we delete the whole paragraph after inserting our new ones
                        bookmark.Parent.Remove();
                    }
                    else
                    {
                        bookmarkData.Element.InsertAfter(bookmark, _document.MainDocumentPart);
                        if (!string.IsNullOrEmpty(bookmarkData.StyleName))
                        {
                            var styleId = GetStyleIdbyName(_document.MainDocumentPart, bookmarkData.StyleName);
                            if (!string.IsNullOrEmpty(styleId))
                            {
                                ((Paragraph)bookmark.Parent).ParagraphProperties =
                                    new ParagraphProperties(new ParagraphStyleId() { Val = styleId });
                            }
                        }
                    }
                }
            }
            else
            {
                Trace.TraceWarning("Bookmark {0} not found", bookmarkData.BookmarkName);
            }
        }

        private Dictionary<string, List<BookmarkStart>> GetAllBookmarks(WordprocessingDocument document)
        {
            var bookmarkMap = new Dictionary<string, List<BookmarkStart>>();

            foreach (BookmarkStart bookmarkStart in document.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                if (!bookmarkMap.ContainsKey(bookmarkStart.Name))
                {
                    bookmarkMap.Add(bookmarkStart.Name, new List<BookmarkStart>());
                }
                bookmarkMap[bookmarkStart.Name].Add(bookmarkStart);
            }
            foreach (var header in document.MainDocumentPart.HeaderParts)
            {
                foreach (BookmarkStart bookmarkStart in header.Header.Descendants<BookmarkStart>())
                {
                    if (!bookmarkMap.ContainsKey(bookmarkStart.Name))
                    {
                        bookmarkMap.Add(bookmarkStart.Name, new List<BookmarkStart>());
                    }
                    bookmarkMap[bookmarkStart.Name].Add(bookmarkStart);
                }
            }
            foreach (var footer in document.MainDocumentPart.FooterParts)
            {
                foreach (BookmarkStart bookmarkStart in footer.Footer.Descendants<BookmarkStart>())
                {
                    if (!bookmarkMap.ContainsKey(bookmarkStart.Name))
                    {
                        bookmarkMap.Add(bookmarkStart.Name, new List<BookmarkStart>());
                    }
                    bookmarkMap[bookmarkStart.Name].Add(bookmarkStart);
                }
            }

            return bookmarkMap;
        }

        private void DeleteBookmarkContent(BookmarkStart bookmarkStart)
        {
            var itemsToDelete = new List<OpenXmlElement>();
            var element = bookmarkStart.NextSibling();
            while (element != null && !(element is BookmarkEnd && ((BookmarkEnd)element).Id.Value == bookmarkStart.Id.Value))
            {
                itemsToDelete.Add(element);
                element = element.NextSibling();
            }

            foreach (var item in itemsToDelete)
            {
                item.Remove();
            }
        }

        private string GetStyleIdbyName(MainDocumentPart mainDocumentPart, string styleName)
        {
            var style = (Style)mainDocumentPart.StyleDefinitionsPart.Styles.FirstOrDefault(s => s is Style && ((Style)s).StyleName.Val.Value.ToLower() == styleName.ToLower());
            return style != null ? style.StyleId.Value : string.Empty;
        }
    }
}
