using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Field = DocumentFormat.OpenXml.Drawing.Field;

namespace SharpDox.Plugins.Word.OpenXml
{
    internal class WordTemplater
    {
        private readonly string _templateFile;

        public WordTemplater(string templateFile)
        {
            _templateFile = templateFile;
        }

        public void ReplaceBookmarks(IEnumerable<FieldData> fieldData)
        {
            foreach (var data in fieldData)
            {
                ReplaceBookmark(data);
            }
        }

        public void ReplaceBookmark(FieldData fieldData)
        {
            using (var document = WordprocessingDocument.Open(_templateFile, true))
            {
                var bookmarks = GetAllBookmarks(document);
                if (bookmarks.ContainsKey(fieldData.FieldName))
                {
                    foreach (var bookmark in bookmarks[fieldData.FieldName])
                    {
                        DeleteBookmarkContent(bookmark);

                        if (fieldData.Element is RichText)
                        {
                            fieldData.Element.InsertAfter(bookmark.Parent,
                                document.MainDocumentPart);
                            // Richtext already includes style information (converted from html), no need to add it

                            // Bookmarks are included in Paragraphs
                            // But paragraphs included in paragraphs are not allowed
                            // Thats why we delete the whole paragraph after inserting our new ones
                            bookmark.Parent.Remove();
                        }
                        else
                        {
                            fieldData.Element.InsertAfter(bookmark, document.MainDocumentPart);
                            if (!string.IsNullOrEmpty(fieldData.StyleName))
                            {
                                var styleId = GetStyleIdbyName(document.MainDocumentPart, fieldData.StyleName);
                                if (!string.IsNullOrEmpty(styleId))
                                {
                                    ((Paragraph)bookmark.Parent).ParagraphProperties =
                                        new ParagraphProperties(new ParagraphStyleId() {Val = styleId});
                                }
                            }
                        }
                    }
                }
                else
                {
                    Trace.TraceWarning("Bookmark {0} not found", fieldData.FieldName);
                }
            }
        }

        public void AddRowToTable(int tableIndex, List<BaseElement> fields)
        {
            using (var document = WordprocessingDocument.Open(_templateFile, true))
            {
                var tables = document.MainDocumentPart.Document.Descendants<Table>().ToList();
                if (tables.Count >= tableIndex + 1)
                {
                    var tableCells = new List<TableCell>();
                    foreach (var field in fields)
                    {
                        var tableCell = new TableCell();
                        var paragraph = new Paragraph();
                        field.AppendTo(paragraph, document.MainDocumentPart);

                        tableCell.Append(paragraph);
                        tableCells.Add(tableCell);
                    }

                    tables[tableIndex].Append(new TableRow(tableCells));
                }
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
            while(element != null && !(element is BookmarkEnd && ((BookmarkEnd)element).Id.Value == bookmarkStart.Id.Value))
            {
                itemsToDelete.Add(element);
                element = element.NextSibling();
            }

            foreach(var item in itemsToDelete)
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
