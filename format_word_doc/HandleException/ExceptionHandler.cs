using System;
using System.Windows;

namespace format_word_doc.HandleException
{
    internal class ExceptionHandler
    {
        public ExceptionHandler(Exception ex)
        {
            MessageBox.Show(ex.Message.ToString());
        }
    }
}
