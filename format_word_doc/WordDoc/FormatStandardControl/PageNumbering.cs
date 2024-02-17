using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class PageNumbering
    {
        public void CreatePageNumber(Word.Document resultDoc)
        {
            foreach (Word.Section wordSection in resultDoc.Sections)
            {
                Word.HeaderFooter footer = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                footer.PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberCenter);
                footer.Range.Font.Name = "Times New Roman";
                footer.Range.Font.Size = 14;
            }
        }
    }
}
