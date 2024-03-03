using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class FormatText
    {
        public void FormattingText(Word.Document resultDoc, Word.Application wordApp)
        {
            Dictionary<string, Action<Word.Paragraph>> titles = new Dictionary<string, Action<Word.Paragraph>>(StringComparer.InvariantCultureIgnoreCase)
            {
                { "ВВЕДЕНИЕ", (paragraph) => { paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; } },
                { "ЗАКЛЮЧЕНИЕ", (paragraph) => { paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; } },
                { "СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ", (paragraph) => { paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; } },
                { "СПИСОК ЛИТЕРАТУРЫ", (paragraph) => { paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; } },
            };

            Word.Range ContentRange = resultDoc.Content;
            Formatting(ContentRange, wordApp, Word.WdParagraphAlignment.wdAlignParagraphThaiJustify);
            
            foreach (var title in titles)
            {
                ContentRange.Find.ClearFormatting();
                ContentRange.Find.Text = title.Key;
                ContentRange.Find.MatchCase = false;
                ContentRange.Find.MatchWholeWord = true;
                ContentRange.Find.MatchWildcards = false;

                while (ContentRange.Find.Execute())
                {
                    Word.Paragraph paragraph = ContentRange.Paragraphs[1];
                    Word.Paragraph paragraphNext = paragraph.Next();
                    title.Value.Invoke(paragraph);
                    ContentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    
                    if (!string.IsNullOrWhiteSpace(paragraphNext.Range.Text))
                    {
                        paragraph.Range.InsertParagraphAfter();
                    }
                }
            }
        }

        public void FormattingAlignmentCenterTitleAutoclaving(Word.Application wordApp, Word.Document resultDoc)
        {
            Word.Range titleAutoclavingRange = resultDoc.Content;
            titleAutoclavingRange.Find.ClearFormatting();
            titleAutoclavingRange.Find.Text = "СОДЕРЖАНИЕ";
            titleAutoclavingRange.Find.Execute();

            if (titleAutoclavingRange.Find.Found)
            {
                Word.Paragraph titleParagraph = titleAutoclavingRange.Paragraphs[1];
                Formatting(titleParagraph, wordApp, Word.WdParagraphAlignment.wdAlignParagraphCenter);

                Word.Paragraph nextParagraph = titleParagraph.Next();
                Formatting(nextParagraph, wordApp, Word.WdParagraphAlignment.wdAlignParagraphCenter);
            }
        }

        public void Formatting(Word.Paragraph paragraph, Word.Application wordApp, Word.WdParagraphAlignment alignment)
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

        private void Formatting(Word.Range range, Word.Application wordApp, Word.WdParagraphAlignment alignment)
        {
            range.Font.Name = "Times New Roman";
            range.Font.Size = 14;
            range.Font.Color = Word.WdColor.wdColorBlack;
            range.ParagraphFormat.Alignment = alignment;
            range.ParagraphFormat.LeftIndent = 0;
            range.ParagraphFormat.RightIndent = 0;
            range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(1.5f);
            range.ParagraphFormat.SpaceBefore = 0;
            range.ParagraphFormat.SpaceAfter = 0;
            range.ParagraphFormat.Space1();
        }
    }
}
