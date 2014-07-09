using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NotesFor.HtmlToOpenXml;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace SharpDox.Plugins.Word.OpenXml
{
    internal class WordTemplater
    {
        private readonly string _templateFile;

        public WordTemplater(string templateFile)
        {
            _templateFile = templateFile;
        }

        public void ReplaceBookmarks(List<FieldData> fieldData)
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
                    DeleteBookmarkContent(bookmarks[fieldData.FieldName]);

                    if (fieldData.IsMarkDown)
                    {
                        var converter = new HtmlConverter(document.MainDocumentPart);
                        var paragraphs = converter.Parse(fieldData.Text);

                        var insertPoint = (OpenXmlElement)bookmarks[fieldData.FieldName].Parent;
                        foreach (var paragraph in paragraphs)
                        {
                            insertPoint.InsertAfterSelf(paragraph);
                            insertPoint = paragraph;
                        }

                        // Bookmarks are included in Paragraphs
                        // But it paragraph included in paragraphs are not allowed
                        // Thats why we delete the whole paragraph after inserting our new ones
                        bookmarks[fieldData.FieldName].Parent.Remove();
                    }
                    else
                    {
                        bookmarks[fieldData.FieldName].InsertAfterSelf(new Run(new Text(fieldData.Text)));
                        if (!string.IsNullOrEmpty(fieldData.StyleName))
                        {
                            var styleId = GetStyleIdbyName(document.MainDocumentPart, fieldData.StyleName);
                            if (!string.IsNullOrEmpty(styleId))
                            {
                                ((Paragraph)bookmarks[fieldData.FieldName].Parent).ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = styleId });
                            }
                        }
                    }
                }
                else
                {
                    Trace.TraceWarning(string.Format("Bookmark {0} not found", fieldData.FieldName));
                }
            }
        }

        public void AddRowToTable(int tableIndex, string[] fields)
        {
            using (var document = WordprocessingDocument.Open(_templateFile, true))
            {
                var tables = document.MainDocumentPart.Document.Descendants<Table>().ToList();
                if (tables.Count >= tableIndex + 1)
                {
                    var tableCells = new List<TableCell>();
                    foreach (var field in fields)
                    {
                        tableCells.Add(new TableCell(new Paragraph(new Run(new Text(field)))));
                    }

                    tables[tableIndex].Append(new TableRow(tableCells));
                }
            }
        }

        private Dictionary<string, BookmarkStart> GetAllBookmarks(WordprocessingDocument document)
        {
            var bookmarkMap = new Dictionary<string, BookmarkStart>();
            foreach (BookmarkStart bookmarkStart in document.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                bookmarkMap[bookmarkStart.Name] = bookmarkStart;
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
