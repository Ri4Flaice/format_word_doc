using format_word_doc.Properties;
using format_word_doc.src.Elements;
using System;
using System.Windows;
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
                menuAnimation.To = 260;
                _isMenuOpen = true;
                menuColumn.Width = new GridLength(260);
            }
            else
            {
                menuAnimation.From = 260;
                menuAnimation.To = 0;
                _isMenuOpen = false;
                menuColumn.Width = new GridLength(40);
            }

            elementsMenuStackPanel.BeginAnimation(StackPanel.WidthProperty, menuAnimation);
        }

        private void SettingsBtn_Click(object sender, RoutedEventArgs e)
        {
            Window parentWindow = Window.GetWindow(this);
            SettingsWindow settingsWindow = new SettingsWindow(parentWindow);
            Opacity = 0.4;
            settingsWindow.ShowDialog();
            Opacity = 1;
        }
        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void SelectAllCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            SetAllCheckBoxesAndSettings(true);
        }

        private void SelectAllCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            SetAllCheckBoxesAndSettings(false);
        }

        private void SetAllCheckBoxesAndSettings(bool isChecked)
        {
            CopyTextCheckBox.IsChecked = isChecked;
            CreateTitlePageCheckBox.IsChecked = isChecked;
            ReplaceTitlePageCheckBox.IsChecked = isChecked;
            CreateHeadingCheckBox.IsChecked = isChecked;
            CreateAutoclavingCheckBox.IsChecked = isChecked;
            FormattingTextCheckBox.IsChecked = isChecked;
            FormattingPictureCheckBox.IsChecked = isChecked;
            PageNumberingCheckBox.IsChecked = isChecked;
            SettingsFieldDocCheckBox.IsChecked = isChecked;

            Settings.Default.CopyTextCheckBox = isChecked;
            Settings.Default.CreateTitlePageCheckBox = isChecked;
            Settings.Default.ReplaceTitlePageCheckBox = isChecked;
            Settings.Default.CreateHeadingCheckBox = isChecked;
            Settings.Default.CreateAutoclavingCheckBox = isChecked;
            Settings.Default.FormattingTextCheckBox = isChecked;
            Settings.Default.FormattingPictureCheckBox = isChecked;
            Settings.Default.PageNumberingCheckBox = isChecked;
            Settings.Default.SettingsFieldDocCheckBox = isChecked;
        }

        private void CopyTextCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.Default.CopyTextCheckBox = true;
        }

        private void CopyTextCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.Default.CopyTextCheckBox = false;
        }

        private void CreateTitlePageCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.Default.CreateTitlePageCheckBox = true;
        }

        private void CreateTitlePageCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.Default.CreateTitlePageCheckBox = false;
        }

        private void ReplaceTitlePageCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.Default.ReplaceTitlePageCheckBox = true;
        }

        private void ReplaceTitlePageCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.Default.ReplaceTitlePageCheckBox = false;
        }

        private void CreateHeadingCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.Default.CreateHeadingCheckBox = true;
        }

        private void CreateHeadingCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.Default.CreateHeadingCheckBox = false;
        }

        private void CreateAutoclavingCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.Default.CreateAutoclavingCheckBox = true;
        }

        private void CreateAutoclavingCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.Default.CreateAutoclavingCheckBox = false;
        }

        private void FormattingTextCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.Default.FormattingTextCheckBox = true;
        }

        private void FormattingTextCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.Default.FormattingTextCheckBox = false;
        }

        private void FormattingPictureCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.Default.FormattingPictureCheckBox = true;
        }

        private void FormattingPictureCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.Default.FormattingPictureCheckBox = false;
        }

        private void SettingsFieldDocCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.Default.SettingsFieldDocCheckBox = true;
        }

        private void SettingsFieldDocCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.Default.SettingsFieldDocCheckBox = false;
        }

        private void PageNumberingCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.Default.PageNumberingCheckBox = true;
        }

        private void PageNumberingCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.Default.PageNumberingCheckBox = false;
        }
    }
}
