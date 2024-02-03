using format_word_doc.src.Elements;
using System.Windows.Controls;

namespace format_word_doc.UserControls
{
    public partial class SettingsUserControl : UserControl
    {
        private BackgroundTextBox _backgroundTextBox;
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
        }

        private void SaveBtn_Click(object sender, System.Windows.RoutedEventArgs e)
        {

        }
    }
}
