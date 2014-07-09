using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;

namespace SharpDox.Plugins.Word.OpenXml
{
    internal static class WordMerger
    {
        public static void AppendPageBreak(string sourceFile)
        {
            using (var document = WordprocessingDocument.Open(sourceFile, true))
            {
                var mainPart = document.MainDocumentPart;
                var para = new Paragraph(new Run((new Break() { Type = BreakValues.Page })));
                mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);
            }
        }

        public static void MergeDocuments(string sourceFile, string destinationFile)
        {
            // I am not using AltChunks, because then all information from the source document get
            // copied to the destination document (including style information etc.). This results
            // in a really big file (filesize wise). To avoid this problem I just copy the body xml and
            // have already included all needed style information in the destination file.
            using (var destinationDocument = WordprocessingDocument.Open(sourceFile, true))
            {
                using(var sourceDocument = WordprocessingDocument.Open(destinationFile, true))
                {
                    // It is necessary to delete the last sectionproperties. Otherwise the resulting document is not valid and Word is not able to open it.
                    sourceDocument.MainDocumentPart.Document.Body.RemoveChild<SectionProperties>(sourceDocument.MainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault());
                    destinationDocument.MainDocumentPart.Document.Body.InnerXml = destinationDocument.MainDocumentPart.Document.Body.InnerXml + sourceDocument.MainDocumentPart.Document.Body.InnerXml;
                    destinationDocument.MainDocumentPart.Document.Save();
                }
            }
        }

        public static void MergeDocumentsWithPagebreak(string sourceFile, string destinationFile)
        {
            AppendPageBreak(sourceFile);
            MergeDocuments(sourceFile, destinationFile);
        }
    }
}
