using RANGER.Database;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Data.SQLite;
using static RANGER.Database.DatabaseImplementation;

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        private void LoginPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            UpdateVisibility();
        }

        private void UpdateVisibility()
        {
            // Ищем TextBlock с watermark в визуальном дереве
            var passwordBox = LoginPasswordBox;
            var grid = VisualTreeHelper.GetChild(passwordBox, 0) as Grid;

            if (grid != null)
            {
                var watermarkText = grid.FindName("watermarkText") as TextBlock;

                if (watermarkText != null)
                {
                    // Показывает плашку "введите ваш пароль..." только если пароль пустой
                    if (string.IsNullOrEmpty(passwordBox.Password))
                    {
                        watermarkText.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        watermarkText.Visibility = Visibility.Collapsed;
                    }
                }
            }
        } 

        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            var DBLoginSession = new DatabaseImplementation().ExecuteQuery<Employee>("SELECT Login, Password FROM Employee;");

            foreach (var Employee in DBLoginSession)
            {
                if (Employee.Employee_ID == 1)
                {
                    MessageBox.Show("Admin conn successful");
                }
                else
                {
                    MessageBox.Show("Kys");
                }
            }
        }
    }
}