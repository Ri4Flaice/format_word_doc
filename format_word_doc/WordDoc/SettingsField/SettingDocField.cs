using Word = Microsoft.Office.Interop.Word;

namespace format_word_doc.WordDoc.SettingsField
{
    internal class SettingDocField
    {
        /// <summary>
        /// margins when creating a document: 
        /// Top = 2 centimeters (0.787402f) 
        /// Left = 3 centimeters (1.1811f) 
        /// Bottom = 2.5 centimeters (0.984252f)
        /// Right = 1 centimeter (0.393701f)
        /// The unit of measurement is inches
        /// </summary>
        /// <param name="wordApp"></param>
        /// <param name="resultDoc"></param>
        public void SettingUpDocumentFields(Word.Application wordApp, Word.Document resultDoc)
        {
            resultDoc.PageSetup.TopMargin = wordApp.InchesToPoints(0.787402f);
            resultDoc.PageSetup.BottomMargin = wordApp.InchesToPoints(0.984252f);
            resultDoc.PageSetup.LeftMargin = wordApp.InchesToPoints(1.1811f);
            resultDoc.PageSetup.RightMargin = wordApp.InchesToPoints(0.393701f);
        }
    }
}
