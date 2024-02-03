using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordManager
{
    internal class WordApplicationManager
    {
        protected Word.Application wordApp = null;
        protected Word.Document titleDoc = null;
        protected Word.Document sourceDoc = null;
        protected Word.Document resultDoc = null;
    }
}
