using System;
using System.IO;
using System.Windows;

namespace format_word_doc.HandleException
{
    internal class ExceptionHandler
    {
        private string logFile = "logs.txt";
        public ExceptionHandler(Exception ex)
        {
            MessageBox.Show(ex.Message.ToString());

            using (StreamWriter writer = new StreamWriter(logFile, true))
            {
                writer.WriteLine("Дата и время ошибки: " + DateTime.Now);
                writer.WriteLine("Сообщение об ошибке: " + ex.Message);
                writer.WriteLine("Стек вызовов: " + ex.StackTrace);
                writer.WriteLine("--------------------------------------------------");
            }
        }
    }
}
