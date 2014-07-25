using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MarkdownSharp;
using SharpDox.Plugins.Word.OpenXml.Elements;
using System.Collections.Generic;
using System.Text;

namespace SharpDox.Plugins.Word.OpenXml
{
    internal class WordTemplater
    {
        private readonly string _templateFile;
        private readonly WordprocessingDocument _document;
        private readonly BookmarkReplacer _bookmarkReplacer;

        public WordTemplater(string templateFile)
        {
            _templateFile = templateFile;
            _document = WordprocessingDocument.Open(_templateFile, true);
            _bookmarkReplacer = new BookmarkReplacer(_document);
        }

        public void ReplaceBookmarks(IEnumerable<BookmarkData> bookmarkData)
        {
            _bookmarkReplacer.ReplaceBookmarks(bookmarkData);
        }

        public void ReplaceBookmark(BookmarkData bookmarkData)
        {
            _bookmarkReplacer.ReplaceBookmark(bookmarkData);
        }

        public void AppendPageBreak()
        {
            var mainPart = _document.MainDocumentPart;
            var para = new Paragraph(new Run((new Break() { Type = BreakValues.Page })));
            mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);
        }

        public void AppendMarkdown(string markdown, string paragraphStyle = null)
        {
            var html = new Markdown().Transform(markdown);
            if (!string.IsNullOrEmpty(paragraphStyle))
            {
                html = html.Replace("<p>", string.Format("<p class=\"{0}\">", paragraphStyle));
            }
            AppendElement(new RichText(html));
        }

        public void AppendImage(string path, string style)
        {
            AppendElement(new RichText(string.Format("<p class=\"{0}\"><img src=\"{1}\"/></p>", style, path)));
        }

        public void AppendParagraph(string text, string style)
        {
            AppendElement(new RichText(string.Format("<p class=\"{0}\">{1}</p>", style, text)));
        }

        public void AppendRichText(string text)
        {
            AppendElement(new RichText(text));
        }

        public void AppendHeader(string text, int navigationlevel)
        {
            AppendElement(new HeaderText(text, navigationlevel));
        }

        public void AppendCodeBlock(string code)
        {
            AppendElement(new RichText(string.Format("<pre><code>{0}</code></pre>", code)));
        }

        public void AppendTable(List<string> headers, List<List<string>> rows, string style = "Table")
        {
            var table = new StringBuilder();
            table.AppendFormat("<table style=\"width:100%\" class=\"{0}\">", style);

            if (headers != null)
            {
                table.Append("<tr>");
                headers.ForEach(h => table.AppendFormat("<th>{0}</th>", h));
                table.Append("</tr>");
            }

            if (rows != null)
            {
                foreach (var row in rows)
                {
                    table.Append("<tr>");
                    row.ForEach(r => table.AppendFormat("<td>{0}</td>", r));
                    table.Append("</tr>");
                }
            }

            table.Append("</table>");

            AppendElement(new RichText(table.ToString()));
        }

        public void AppendList(List<string> elements)
        {
            var list = new StringBuilder();
            list.Append("<ul>");
            elements.ForEach(e => list.AppendFormat("<li>{0}</li>", e));
            list.Append("</ul>");

            AppendElement(new RichText(list.ToString()));
        }

        public void Close()
        {
            _document.Close();
        }

        private void AppendElement(BaseElement element)
        {
            element.AppendTo(_document.MainDocumentPart.Document.Body, _document.MainDocumentPart);
        }
    }
}
