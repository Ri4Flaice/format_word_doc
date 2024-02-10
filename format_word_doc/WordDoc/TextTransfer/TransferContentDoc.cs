using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.TextTransfer
{
    internal class TransferContentDoc
    {
        public void TransferringContentFromOriginalToNewDoc(Word.Document sourceDoc, Word.Application wordApp, Word.Document resultDoc)
        {
            object start = resultDoc.Content.End - 1;
            object end = resultDoc.Content.End;
            Word.Range range = resultDoc.Range(ref start, ref end);
            range.InsertFile(sourceDoc.FullName);
        }
    }
}
