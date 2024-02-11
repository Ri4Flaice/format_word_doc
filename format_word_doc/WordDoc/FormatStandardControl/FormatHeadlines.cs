using System.Collections.Generic;
using System;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class FormatHeadlines
    {
        private Word.Paragraph _previousParagraphBeforeTitle = null;

        /// <summary>
        /// Add functionality: find the word APPLICATION and make heading 1 for it
        /// </summary>
        /// <param name="resultDoc"></param>
        public void FindTitleInText(Word.Document resultDoc)
        {
            Dictionary<string, Action<Word.Paragraph>> keyWords = new Dictionary<string, Action<Word.Paragraph>>(StringComparer.InvariantCultureIgnoreCase)
            {
                { "ВВЕДЕНИЕ", (paragraph) => SetStyleHeading(paragraph, Word.WdBuiltinStyle.wdStyleHeading1) },
                { "ЗАКЛЮЧЕНИЕ", (paragraph) => { SetStyleHeading(paragraph, Word.WdBuiltinStyle.wdStyleHeading1); AddPageBreakBeforeTitle(); } },
                { "СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ", (paragraph) => { SetStyleHeading(paragraph, Word.WdBuiltinStyle.wdStyleHeading1); AddPageBreakBeforeTitle(); } },
                { "СПИСОК ЛИТЕРАТУРЫ", (paragraph) => { SetStyleHeading(paragraph, Word.WdBuiltinStyle.wdStyleHeading1); AddPageBreakBeforeTitle(); } }
            };

            foreach (Word.Paragraph paragraph in resultDoc.Paragraphs)
            {
                if (keyWords.ContainsKey(paragraph.Range.Text.Trim()))
                {
                    keyWords[paragraph.Range.Text.Trim()].Invoke(paragraph);
                }
                else if (Regex.IsMatch(paragraph.Range.Text, @"^\d\s"))
                {
                    SetStyleHeading(paragraph, Word.WdBuiltinStyle.wdStyleHeading1);
                    AddPageBreakBeforeTitle();
                }
                else if (Regex.IsMatch(paragraph.Range.Text, @"^\d\.\d\s"))
                {
                    SetStyleHeading(paragraph, Word.WdBuiltinStyle.wdStyleHeading2);
                }

                _previousParagraphBeforeTitle = paragraph;
            }
        }

        private void SetStyleHeading(Word.Paragraph paragraph, Word.WdBuiltinStyle styleHeading)
        {
            paragraph.set_Style(styleHeading);
        }

        private void AddPageBreakBeforeTitle()
        {
            Word.Range range = _previousParagraphBeforeTitle.Range;
            range.InsertAfter("\f");
        }
    }
}
