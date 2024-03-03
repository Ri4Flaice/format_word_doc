using format_word_doc.HandleException;
using System;
using System.IO;

namespace format_word_doc.src.CreateDirecoty
{
    internal class DirectoryDocuments
    {
        public void CreateDirectoryDocuments()
        {
            try
            {
                string exeDirectoryPath = AppDomain.CurrentDomain.BaseDirectory;
                string documentsDirectoryPath = Path.Combine(exeDirectoryPath, "Documents");

                if (!Directory.Exists(documentsDirectoryPath))
                {
                    Directory.CreateDirectory(documentsDirectoryPath);
                }
            }
            catch (Exception ex)
            {
                new ExceptionHandler(ex);
            }
        }
    }
}
