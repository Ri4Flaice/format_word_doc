using format_word_doc.HandleException;
using format_word_doc.WordManager;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.CreateDocument
{
    internal class CreateDoc : WordApplicationManager
    {
        public void CreateDocument()
        {
            try
            {
                wordApp = new Word.Application();
                resultDoc = wordApp.Documents.Add();

                string exeFolderPath = AppDomain.CurrentDomain.BaseDirectory;
                string fullPath = System.IO.Path.Combine(exeFolderPath, "Documents/result.docx");
                
                resultDoc.SaveAs(fullPath);
            }
            catch (Exception ex)
            {
                new ExceptionHandler(ex);
            }
            finally
            {
                resultDoc.Close();
                wordApp.Quit();
            }
        }
    }
}
