using System.Collections.Generic;
using System;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace format_word_doc.WordDoc.FormatStandardControl
{
    internal class FormatText
    {
        public void FormattingText(Word.Document resultDoc, Word.Application wordApp, byte startPage = 0)
        {
            Dictionary<string, Action<Word.Paragraph>> keyWords = new Dictionary<string, Action<Word.Paragraph>>(StringComparer.InvariantCultureIgnoreCase)
            {
                { "ВВЕДЕНИЕ", (paragraph) => { paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; } },
                { "ЗАКЛЮЧЕНИЕ", (paragraph) => { paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; } },
                { "СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ", (paragraph) => { paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; } },
                { "СПИСОК ЛИТЕРАТУРЫ", (paragraph) => { paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; } },
            };

            foreach (Word.Paragraph paragraph in resultDoc.Paragraphs)
            {
                if (paragraph.Range.Information[Word.WdInformation.wdActiveEndPageNumber] > startPage)
                {
                    Formatting(paragraph, wordApp, Word.WdParagraphAlignment.wdAlignParagraphJustify);

                    if (Regex.IsMatch(paragraph.Range.Text, @"^приложение\s*\w", RegexOptions.IgnoreCase))
                    {
                        if (paragraph.Range.Information[Word.WdInformation.wdActiveEndPageNumber] > 3)
                        {
                            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                    }
                    else if (Regex.IsMatch(paragraph.Range.Text, @"^продолжение приложения\s*\w", RegexOptions.IgnoreCase))
                    {
                        paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    }
                }
                if (keyWords.ContainsKey(paragraph.Range.Text.Trim()))
                {
                    keyWords[paragraph.Range.Text.Trim()].Invoke(paragraph);
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
    }
}
