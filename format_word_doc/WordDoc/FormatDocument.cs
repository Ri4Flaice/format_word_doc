using format_word_doc.HandleException;
using format_word_doc.Properties;
using format_word_doc.WordDoc.FormatStandardControl;
using format_word_doc.WordDoc.SettingsField;
using format_word_doc.WordDoc.TextTransfer;
using format_word_doc.WordDoc.TitlePage;
using format_word_doc.WordManager;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc
{
    internal class FormatDocument : WordApplicationManager
    {
        private WorkTitle _workTitle = new WorkTitle();
        private TransferContentDoc _transferContentDoc = new TransferContentDoc();
        private FormatText _formatText = new FormatText();
        private FormatPicture _formatPicture = new FormatPicture();
        private SettingDocField _settingDocField = new SettingDocField();
        public void Formatting()
        {
            try
            {
                wordApp = new Word.Application();

                string exeFolderPath = AppDomain.CurrentDomain.BaseDirectory;

                string titleDocPath = System.IO.Path.Combine(exeFolderPath, "Documents/Title.docx");
                string sourceDocPath = System.IO.Path.Combine(exeFolderPath, "Documents/source.docx");
                string resultDocPath = System.IO.Path.Combine(exeFolderPath, "Documents/result.docx");

                titleDoc = wordApp.Documents.Open(titleDocPath);
                sourceDoc = wordApp.Documents.Open(sourceDocPath);
                resultDoc = wordApp.Documents.Open(resultDocPath);

                if (Settings.Default.CopyTextCheckBox) { _transferContentDoc.TransferringContentFromOriginalToNewDoc(sourceDoc, wordApp, resultDoc); }
                if (Settings.Default.CreateTitlePageCheckBox) { _workTitle.CopyTitleOfTheTitleDoc(resultDoc, titleDocPath); _workTitle.ReplaceContentTitlePage(resultDoc); }
                if (Settings.Default.FormattingTextCheckBox) { _formatText.FormattingText(resultDoc, wordApp, 1); }
                if (Settings.Default.FormattingPictureCheckBox) { _formatPicture.FormattingPicture(resultDoc); }
                if (Settings.Default.SettingsFieldDocCheckBox) { _settingDocField.SettingUpDocumentFields(wordApp, resultDoc); }

                resultDoc.Save();
            }
            catch (Exception ex)
            {
                new ExceptionHandler(ex);
            }
            finally
            {
                titleDoc.Close();
                sourceDoc.Close();
                resultDoc.Close();
                wordApp.Quit();
            }
        }
    }
}
