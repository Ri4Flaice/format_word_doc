using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class FormatTable
    {
        public void FormattingCellsAlignmentCenter(Word.Application wordApp, Word.Document resultDoc)
        {
            wordApp.ScreenUpdating = false;

            for (int i = 1; i <= resultDoc.Tables.Count; i++)
            {
                Word.Table table = resultDoc.Tables[i];
                
                table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                table.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);

                AddEmptyParagraphNearTable(table, true);
                AddEmptyParagraphNearTable(table, false);
            }

            wordApp.ScreenUpdating = true;
        }

        private void AddEmptyParagraphNearTable(Word.Table table, bool tableRangeStart)
        {
            Word.Range rangeTable = table.Range.Duplicate;

            if (tableRangeStart)
            {
                rangeTable.SetRange(table.Range.Start - 1, table.Range.Start - 1);
            }
            else
            {
                rangeTable.SetRange(table.Range.End, table.Range.End);
            }

            rangeTable.Text = "\r";
        }
    }
}
