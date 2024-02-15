using format_word_doc.Properties;
using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class Autoclaving
    {
        public void CreateAutoclaving(Word.Document resultDoc)
        {
            const string AUTOCLAVINGTITLE = "СОДЕРЖАНИЕ";
            Word.Range endOfTitlePage;

            if (Settings.Default.CreateTitlePageCheckBox)
            {
                endOfTitlePage = StartPositionAutoclaving(resultDoc, 2);
            }
            else
            {
                endOfTitlePage = StartPositionAutoclaving(resultDoc, 1);
            }

            endOfTitlePage.InsertAfter(AUTOCLAVINGTITLE + "\n");
            endOfTitlePage.set_Style(Word.WdBuiltinStyle.wdStyleNormal);

            Word.Range startOfAutoclaving = resultDoc.Range(endOfTitlePage.End, endOfTitlePage.End);
            Word.TableOfContents autoclaving = resultDoc.TablesOfContents.Add(startOfAutoclaving, UseHyperlinks: true, UseOutlineLevels: true);
            Word.Range autoclavingRange = autoclaving.Range;
            Word.Range endOfAutoclaving = resultDoc.Range(autoclavingRange.End, autoclavingRange.End);

            endOfAutoclaving.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
        }

        private Word.Range StartPositionAutoclaving(Word.Document resultDoc, int numberPage)
        {
            return resultDoc.GoTo(What: Word.WdGoToItem.wdGoToPage, Which: Word.WdGoToDirection.wdGoToAbsolute, Count: numberPage);
        }
    }
}
