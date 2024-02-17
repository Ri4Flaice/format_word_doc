using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class PageNumbering
    {
        public void CreatePageNumber(Word.Document resultDoc, int startNumberPage)
        {
            foreach (Word.Section wordSection in resultDoc.Sections)
            {
                Word.HeaderFooter footer = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                footer.LinkToPrevious = false;

                if (wordSection.Index > startNumberPage)
                {
                    footer.PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberCenter);
                    footer.Range.Font.Name = "Times New Roman";
                    footer.Range.Font.Size = 14;
                }
                else
                {
                    footer.Range.Delete();
                }
            }
        }
    }
}
