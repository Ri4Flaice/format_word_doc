using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class FormatPicture
    {
        public void FormattingPicture(Word.Document resultDoc)
        {
            foreach (Word.InlineShape picture in resultDoc.InlineShapes)
            {
                if (picture.Type == Word.WdInlineShapeType.wdInlineShapePicture)
                {
                    CheckEmptyParagraphBeforePicture(picture, resultDoc);

                    picture.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    picture.Line.ForeColor.RGB = (int)Word.WdColor.wdColorBlack;
                    picture.Line.Weight = 1.0f;

                    CheckEmptyParagraphAfterPicture(picture, resultDoc);
                }
            }
        }

        private void CheckEmptyParagraphBeforePicture(Word.InlineShape picture, Word.Document resultDoc)
        {
            if (picture.Range.Paragraphs.First.Previous().Range.Text != "\r")
            {
                Word.Paragraph paragraph = resultDoc.Content.Paragraphs.Add(picture.Range);
                paragraph.Range.InsertBefore("\n");
            }
        }

        private void CheckEmptyParagraphAfterPicture(Word.InlineShape picture, Word.Document resultDoc)
        {
            if (picture.Range.Paragraphs.First.Next().Range.Text != "\r")
            {
                Word.Paragraph paragraph = resultDoc.Content.Paragraphs.Add(picture.Range);
                paragraph.Range.InsertAfter("\n");
            }
        }
    }
}
