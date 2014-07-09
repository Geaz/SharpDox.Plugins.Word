using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;

namespace SharpDox.Plugins.Word.OpenXml
{
    internal static class WordMerger
    {
        private static int _chunkId = 1;

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
            using (var document = WordprocessingDocument.Open(destinationFile, true))
            {
                var altChunkId = string.Format("AltChunkId{0}", _chunkId++);
                var mainPart = document.MainDocumentPart;
                AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

                using (FileStream fileStream = File.Open(sourceFile, FileMode.Open))
                {
                    chunk.FeedData(fileStream);
                }

                var altChunk = new AltChunk { Id = altChunkId };
                mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements().Last());
            }
        }

        public static void MergeDocumentsWithPagebreak(string sourceFile, string destinationFile)
        {
            AppendPageBreak(destinationFile);
            MergeDocuments(sourceFile, destinationFile);
        }
    }
}
