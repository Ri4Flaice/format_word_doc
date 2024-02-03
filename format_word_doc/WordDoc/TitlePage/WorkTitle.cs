using format_word_doc.Properties;
using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.TitlePage
{
    internal class WorkTitle
    {
        private Dictionary<string, string> _defaultSettings = new Dictionary<string, string>
        {
            { "<NameMinistryEducation>", Settings.Default.MinistryEducation },
            { "<NameOrganization>", Settings.Default.Organization },
            { "<NameFaculty>", Settings.Default.Faculty },
            { "<NameDepartment>", Settings.Default.Department },
            { "<Theme>", Settings.Default.Theme },
            { "<NameStudent>", Settings.Default.Student },
            { "<GenderStudent>", Settings.Default.Completed },
            { "<City>", Settings.Default.City },
            { "<NameGroup>", Settings.Default.Group },
            { "<NameTeacher>", Settings.Default.Teacher },
            { "<CheckedTeacher>", Settings.Default.Checked },
            { "<PostTeacher>", Settings.Default.Post },
            { "<CurrentYear>", DateTime.Now.Year.ToString() }
        };

        public void CopyTitleOfTheTitleDoc(Word.Document resultDoc, string pathTitleDoc)
        {
            Word.Paragraph firstParagraph = resultDoc.Paragraphs[1];
            firstParagraph.Range.InsertParagraphBefore();
            firstParagraph.Previous().Range.InsertFile(pathTitleDoc);

            Word.Range rangeAfterInsertedText = resultDoc.Range(firstParagraph.Previous().Range.End, firstParagraph.Previous().Range.End);
            rangeAfterInsertedText.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
        }

        public void ReplaceContentTitlePage(Word.Document resultDoc)
        {
            foreach (var setting in _defaultSettings)
            {
                var find = resultDoc.Content.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = setting.Key;
                find.Replacement.Text = setting.Value;

                find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }
        }
    }
}
