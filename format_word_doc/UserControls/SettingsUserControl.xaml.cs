using format_word_doc.Properties;
using format_word_doc.src.Elements;
using System.Windows;
using System;
using System.Windows.Controls;
using format_word_doc.HandleException;

namespace format_word_doc.UserControls
{
    public partial class SettingsUserControl : UserControl
    {
        private BackgroundTextBox _backgroundTextBox;
        private TextBox[] _textBoxes;
        private string[] _settingsTextBoxes;
        public SettingsUserControl()
        {
            InitializeComponent();

            _backgroundTextBox = new BackgroundTextBox();

            _backgroundTextBox.SetTextBoxPlaceholder(ministryEducationTextBox, "Введите полное название министерства");
            _backgroundTextBox.SetTextBoxPlaceholder(organizationTextBox, "Введите полное название университета");
            _backgroundTextBox.SetTextBoxPlaceholder(facultyTextBox, "Введите полное название факультета");
            _backgroundTextBox.SetTextBoxPlaceholder(departmentTextBox, "Введите полное название кафедры");
            _backgroundTextBox.SetTextBoxPlaceholder(themeTextBox, "Введите тему");
            _backgroundTextBox.SetTextBoxPlaceholder(nameStudentTextBox, "Введите ФИО (Иванов И.И.)");
            _backgroundTextBox.SetTextBoxPlaceholder(completedTextBox, "Введите (Выполнил(-а) студент)");
            _backgroundTextBox.SetTextBoxPlaceholder(groupTextBox, "Введите инициалы группы");
            _backgroundTextBox.SetTextBoxPlaceholder(cityTextBox, "Введите название города");
            _backgroundTextBox.SetTextBoxPlaceholder(nameTeacherTextBox, "Введите ФИО преподавателя (Иванова И.И.)");
            _backgroundTextBox.SetTextBoxPlaceholder(checkedTeacherTextBox, "Введите (Проверил(-а))");
            _backgroundTextBox.SetTextBoxPlaceholder(postTeacherTextBox, "Введите должность преподавателя");

            InitializeControls();
            InitializeSettings();

            AssigningDataTextBox();
        }

        private void InitializeControls()
        {
            _textBoxes = new TextBox[]
            {
                themeTextBox, ministryEducationTextBox, organizationTextBox,
                facultyTextBox, departmentTextBox, nameStudentTextBox, completedTextBox,
                groupTextBox, cityTextBox, nameTeacherTextBox, checkedTeacherTextBox,
                postTeacherTextBox,
            };
        }

        private void InitializeSettings()
        {
            _settingsTextBoxes = new string[]
            {
                Settings.Default.Theme, Settings.Default.MinistryEducation,
                Settings.Default.Organization, Settings.Default.Faculty, Settings.Default.Department,
                Settings.Default.Student, Settings.Default.Completed, Settings.Default.Group,
                Settings.Default.City, Settings.Default.Teacher, Settings.Default.Checked,
                Settings.Default.Post
            };
        }

        private void AssigningDataTextBox()
        {
            try
            {
                for (int i = 0; i < _textBoxes.Length; i++)
                {
                    _textBoxes[i].Text = _settingsTextBoxes[i];
                }
            }
            catch (Exception ex)
            {
                new ExceptionHandler(ex);
            }
        }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Settings.Default.Save();
                System.Diagnostics.Process.Start(Application.ResourceAssembly.Location);
                Application.Current.Shutdown();

            }
            catch (Exception ex)
            {
                new ExceptionHandler(ex);
            }
        }

        private void ministryEducationTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.MinistryEducation = ministryEducationTextBox.Text;
        }

        private void organizationTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Organization = organizationTextBox.Text;
        }

        private void facultyTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Faculty = facultyTextBox.Text;
        }
        private void departmentTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Department = departmentTextBox.Text;
        }

        private void themeTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Theme = themeTextBox.Text;
        }

        private void nameStudentTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Student = nameStudentTextBox.Text;
        }

        private void completedTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Completed = completedTextBox.Text;
        }

        private void groupTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Group = groupTextBox.Text;
        }

        private void cityTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.City = cityTextBox.Text;
        }

        private void nameTeacherTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Teacher = nameTeacherTextBox.Text;
        }

        private void checkedTeacherTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Checked = checkedTeacherTextBox.Text;
        }

        private void postTeacherTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Settings.Default.Post = postTeacherTextBox.Text;
        }
    }
}
