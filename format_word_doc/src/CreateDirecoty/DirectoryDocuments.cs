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
                string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string documentsDirectory = Path.Combine(exeDirectory, "Documents");

                if (!Directory.Exists(documentsDirectory))
                {
                    Directory.CreateDirectory(documentsDirectory);
                }
            }
            catch (Exception ex)
            {
                new ExceptionHandler(ex);
            }
        }
    }
}
