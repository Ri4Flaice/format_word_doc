using format_word_doc.HandleException;
using format_word_doc.WordDoc.FormatStandardControl;
using format_word_doc.WordManager;
using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.CreateDocument
{
    internal class CreateTitleDoc : WordApplicationManager
    {
        private FormatText _formatText = new FormatText();
        public void CreateTitleDocument()
        {
            try
            {
                string exeDirectoryPath = AppDomain.CurrentDomain.BaseDirectory;
                string titleDocumentPath = Path.Combine(exeDirectoryPath, "Documents/title.docx");

                if (!File.Exists(titleDocumentPath))
                {
                    wordApp = new Word.Application();
                    titleDoc = wordApp.Documents.Add();

                    Word.Range range = titleDoc.Range();

                    AddTemplateTexts(range);
                    Word.Table table = AddTemplateTable(range);
                    DeleteFirstParagraph();

                    _formatText.Formatting(titleDoc.Content, wordApp, Word.WdParagraphAlignment.wdAlignParagraphCenter);

                    AlignmentCellsTable(table);

                    titleDoc.SaveAs2(titleDocumentPath);
                }
            }
            catch (Exception ex)
            {
                new ExceptionHandler(ex);
            }
            finally
            {
                titleDoc.Close();
                wordApp.Quit();
            }
        }

        /// <summary>
        /// We add text according to the template so that there are no differences
        /// </summary>
        /// <param name="range"></param>
        private void AddTemplateTexts(Word.Range range)
        {
            range.Text += "<NameMinistryEducation>";
            range.Text += "<NameOrganization>";
            range.Text += "<NameFaculty>";
            range.Text += "<NameDepartment>\n\n\n\n\n\n\n\n";
            range.Text += "<Theme>\n\n\n\n\n\n\n\n";
            range.Text += "\n\n\n\n\n\n\n\n\n\n\n\n<City>, <CurrentYear>";
        }

        private Word.Table AddTemplateTable(Word.Range range)
        {
            Word.Paragraph tableParagraph = titleDoc.Paragraphs[21];
            tableParagraph.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.End = tableParagraph.Range.End;
            Word.Table table = titleDoc.Tables.Add(range, 2, 2);

            table.PreferredWidth = wordApp.CentimetersToPoints(15.92f);
            table.Columns[1].PreferredWidth = wordApp.CentimetersToPoints(7.96f);
            table.Columns[2].PreferredWidth = wordApp.CentimetersToPoints(7.96f);

            table.BottomPadding = wordApp.CentimetersToPoints(0.18f);
            table.LeftPadding = wordApp.CentimetersToPoints(0.18f);
            table.TopPadding = wordApp.CentimetersToPoints(0.18f);
            table.RightPadding = wordApp.CentimetersToPoints(0.18f);

            table.Cell(1, 1).Range.Text = "<GenderStudent>\n<NameGroup>";
            table.Cell(2, 1).Range.Text = "<CheckedTeacher>\n<PostTeacher>";
            table.Cell(1, 2).Range.Text = "<NameStudent>";
            table.Cell(2, 2).Range.Text = "<NameTeacher>";

            table.Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            return table;
        }

        private void AlignmentCellsTable(Word.Table table)
        {
            table.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;

            table.Cell(2, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(2, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;

            table.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            table.Cell(2, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(2, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        }

        /// <summary>
        /// When adding a paragraph to the beginning of the document, the result is an empty paragraph, 
        /// so we delete it
        /// </summary>
        private void DeleteFirstParagraph()
        {
            Word.Paragraph firstParagraph = titleDoc.Paragraphs.First;
            firstParagraph.Range.Delete();
        }
    }
}
