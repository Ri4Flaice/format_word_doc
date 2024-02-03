using format_word_doc.HandleException;
using format_word_doc.WordManager;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc
{
    internal class FormatDocument : WordApplicationManager
    {
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
