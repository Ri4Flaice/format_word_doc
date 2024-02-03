using System.Windows;

namespace format_word_doc
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow(Window parentWindow)
        {
            Owner = parentWindow;
            InitializeComponent();
        }
    }
}
