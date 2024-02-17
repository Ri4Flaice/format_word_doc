using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class PageNumbering
    {
        public void CreatePageNumber(Word.Document resultDoc)
        {
            Word.Section firstSection = resultDoc.Sections[1];
            Word.HeaderFooter firstFooter = firstSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            firstFooter.LinkToPrevious = false;
            firstFooter.Range.Delete();

            for (int i = 2; i <= resultDoc.Sections.Count; i++)
            {
                Word.Section wordSection = resultDoc.Sections[i];
                Word.HeaderFooter footer = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                footer.PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberCenter);
                footer.Range.Font.Name = "Times New Roman";
                footer.Range.Font.Size = 14;
            }

            //foreach (Word.Section wordSection in resultDoc.Sections)
            //{
            //    Word.HeaderFooter footer = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            //    footer.PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberCenter);
            //    footer.Range.Font.Name = "Times New Roman";
            //    footer.Range.Font.Size = 14;
            //}
        }
    }
}
