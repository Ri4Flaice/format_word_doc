using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class FormatText
    {
        public void FormattingText(Word.Document resultDoc, Word.Application wordApp, byte startPage = 0)
        {
            foreach (Word.Paragraph paragraph in resultDoc.Paragraphs)
            {
                if (paragraph.Range.Information[Word.WdInformation.wdActiveEndPageNumber] > startPage)
                {
                    Formatting(paragraph, wordApp, Word.WdParagraphAlignment.wdAlignParagraphJustify);
                }
            }
        }

        private void Formatting(Word.Paragraph paragraph, Word.Application wordApp, Word.WdParagraphAlignment alignment)
        {
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Size = 14;
            paragraph.Range.Font.Color = Word.WdColor.wdColorBlack;
            paragraph.Alignment = alignment;
            paragraph.Format.LeftIndent = 0;
            paragraph.Format.RightIndent = 0;
            paragraph.Format.FirstLineIndent = wordApp.CentimetersToPoints(1.5f);
            paragraph.Format.SpaceBefore = 0;
            paragraph.Format.SpaceAfter = 0;
            paragraph.Space1();
        }
    }
}
