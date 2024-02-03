using System.Windows.Controls;
using System.Windows.Media;

namespace format_word_doc.src.Elements
{
    internal class BackgroundTextBox
    {
        public void SetTextBoxPlaceholder(TextBox textBox, string placeholderText)
        {
            textBox.Background = new VisualBrush()
            {
                Visual = new Label() { Content = placeholderText, Foreground = Brushes.Gray },
                Opacity = 0.6,
                Stretch = Stretch.None,
                AlignmentX = AlignmentX.Left,
                AlignmentY = AlignmentY.Center
            };
            textBox.TextChanged += (s, e) =>
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                    textBox.Background = Brushes.Transparent;
                else
                    textBox.Background = new VisualBrush()
                    {
                        Visual = new Label() { Content = placeholderText, Foreground = Brushes.Gray },
                        Opacity = 0.6,
                        Stretch = Stretch.None,
                        AlignmentX = AlignmentX.Left,
                        AlignmentY = AlignmentY.Center
                    };
            };
        }
    }
}
