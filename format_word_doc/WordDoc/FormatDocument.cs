using format_word_doc.HandleException;
using format_word_doc.Properties;
using format_word_doc.WordDoc.FormatStandardControl;
using format_word_doc.WordDoc.SettingsField;
using format_word_doc.WordDoc.TextTransfer;
using format_word_doc.WordDoc.TitlePage;
using format_word_doc.WordManager;
using System;
using System.IO;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc
{
    internal class FormatDocument : WordApplicationManager
    {
        private WorkTitle _workTitle = new WorkTitle();
        private TransferContentDoc _transferContentDoc = new TransferContentDoc();
        private FormatHeadlines _formatHeadlines = new FormatHeadlines();
        private Autoclaving _autoclaving = new Autoclaving();
        private FormatText _formatText = new FormatText();
        private FormatTable _formatTable = new FormatTable();
        private FormatPicture _formatPicture = new FormatPicture();
        private SettingDocField _settingDocField = new SettingDocField();
        private PageNumbering _pageNumbering = new PageNumbering();
        public void Formatting()
        {
            try
            {
                wordApp = new Word.Application();

                string exeFolderPath = AppDomain.CurrentDomain.BaseDirectory;
                int startNumberPage = 0;

                string titleDocPath = Path.Combine(exeFolderPath, "Documents\\Title.docx");
                string resultDocPath = Path.Combine(exeFolderPath, "Documents\\result.docx");
                string sourceDocPath = SourceFilePath(exeFolderPath, titleDocPath, resultDocPath);

                titleDoc = wordApp.Documents.Open(titleDocPath);
                sourceDoc = wordApp.Documents.Open(sourceDocPath);
                resultDoc = wordApp.Documents.Open(resultDocPath);

                if (Settings.Default.CopyTextCheckBox) { _transferContentDoc.TransferringContentFromOriginalToNewDoc(sourceDoc, wordApp, resultDoc); }
                if (Settings.Default.CreateHeadingCheckBox) { _formatHeadlines.FindTitleInText(resultDoc); }
                
                if (Settings.Default.FormattingTextCheckBox) { 
                    _formatText.FormattingText(resultDoc, wordApp); 
                    _formatTable.FormattingCellsAlignmentCenter(wordApp, resultDoc);
                    _formatHeadlines.AlignmentCenterWordsApplication(resultDoc);
                }
                
                if (Settings.Default.CreateTitlePageCheckBox) { _workTitle.CopyTitleOfTheTitleDoc(resultDoc, titleDocPath); _workTitle.ReplaceContentTitlePage(resultDoc); startNumberPage++; }
                
                if (Settings.Default.FormattingPictureCheckBox) { _formatPicture.FormattingPicture(resultDoc); }
                
                if (Settings.Default.SettingsFieldDocCheckBox) { _settingDocField.SettingUpDocumentFields(wordApp, resultDoc); }
                
                if (Settings.Default.CreateAutoclavingCheckBox) { _autoclaving.CreateAutoclaving(wordApp, resultDoc); startNumberPage++; }
                if (Settings.Default.CreateAutoclavingCheckBox) { _formatText.FormattingAlignmentCenterTitleAutoclaving(wordApp, resultDoc); }
                
                if (Settings.Default.PageNumberingCheckBox) { _pageNumbering.CreatePageNumber(resultDoc, startNumberPage); }

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

        private string SourceFilePath(string exeDirectoryPath, string titleDocumentPath, string resultDocumentPath)
        {
            string[] files = Directory.GetFiles(Path.Combine(exeDirectoryPath, "Documents"));
            string sourceDocumentPath = null;

            if (files.Length != 3)
            {
                MessageBox.Show("В директории должно быть ровно три файла\nДля корректной работы");
                Environment.Exit(0);
            }

            foreach (string file in files)
            {
                if (file != titleDocumentPath && file != resultDocumentPath)
                {
                    sourceDocumentPath = file;
                    break;
                }
            }

            return sourceDocumentPath;
        }
    }
}
