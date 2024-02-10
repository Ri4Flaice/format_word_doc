using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.TextTransfer
{
    internal class TransferContentDoc
    {
        public void TransferringContentFromOriginalToNewDoc(Word.Document sourceDoc, Word.Application wordApp, Word.Document resultDoc)
        {
            sourceDoc.Content.Select();
            wordApp.Selection.Copy();
            resultDoc.Content.Paste();
        }
    }
}
