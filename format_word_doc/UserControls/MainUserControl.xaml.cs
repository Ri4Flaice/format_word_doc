using System.Windows;
using System;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace format_word_doc.UserControls
{
    public partial class MainUserControl : UserControl
    {
        private bool _isMenuOpen = false;
        public MainUserControl()
        {
            InitializeComponent();
        }

        private void OpenClosedMenuBtn_Click(object sender, RoutedEventArgs e)
        {
            DoubleAnimation menuAnimation = new DoubleAnimation();
            menuAnimation.Duration = TimeSpan.FromMilliseconds(50);

            if (!_isMenuOpen)
            {
                menuAnimation.From = elementsMenuStackPanel.ActualWidth;
                menuAnimation.To = 230;
                _isMenuOpen = true;
                menuColumn.Width = new GridLength(230);
            }
            else
            {
                menuAnimation.From = 230;
                menuAnimation.To = 0;
                _isMenuOpen = false;
                menuColumn.Width = new GridLength(40);
            }

            elementsMenuStackPanel.BeginAnimation(StackPanel.WidthProperty, menuAnimation);
        }

        private void SettingsBtn_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void SelectAllCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            CopyTextCheckBox.IsChecked = true;
            CreateTitlePageCheckBox.IsChecked = true;
            ReplaceTitlePageCheckBox.IsChecked = true;
            CreateHeadingCheckBox.IsChecked = true;
            CreateAutoclavingCheckBox.IsChecked = true;
            FormattingTextCheckBox.IsChecked = true;
            FormattingPictureCheckBox.IsChecked = true;
            PageNumberingCheckBox.IsChecked = true;
            SettingsFieldDocCheckBox.IsChecked = true;
        }

        private void SelectAllCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            CopyTextCheckBox.IsChecked = false;
            CreateTitlePageCheckBox.IsChecked = false;
            ReplaceTitlePageCheckBox.IsChecked = false;
            CreateHeadingCheckBox.IsChecked = false;
            CreateAutoclavingCheckBox.IsChecked = false;
            FormattingTextCheckBox.IsChecked = false;
            FormattingPictureCheckBox.IsChecked = false;
            PageNumberingCheckBox.IsChecked = false;
            SettingsFieldDocCheckBox.IsChecked = false;
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
