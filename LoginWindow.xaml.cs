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
using static RANGER.Database.DataBaseConnection;
using static RANGER.Database.DataBaseModel;

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
            UpdateVisibilityForPasswordBox();
        }

        private void UpdateVisibilityForPasswordBox()
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

        private void LoginTextBox_LoginChanged(object sender, RoutedEventArgs e)
        {
            // То же самое что и у PasswordBox, но для TextBlock для этого элемента
            var loginTextBox = LoginTextBox;
            var grid = VisualTreeHelper.GetChild(loginTextBox, 0) as Grid;

            var watermarkText = grid.FindName("watermarkText") as TextBlock;

            if(watermarkText != null)
            {
                if (string.IsNullOrEmpty(loginTextBox.Text))
                {
                    watermarkText.Visibility = Visibility.Visible;
                }
                else
                {
                    watermarkText.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            var DBLoginSession = new DataBaseConnection().ExecuteQuery<Employee>("SELECT * FROM Employee;");

            var user = DBLoginSession.Where(n => n.Password == LoginPasswordBox.Password && n.Login == LoginTextBox.Text).FirstOrDefault();

            if (user != null)
            {
                MainWindow mainWindow = new MainWindow(user);
                mainWindow.Show();
                Hide();
            }
            else
            {
                MessageBox.Show("Неверный логин или пароль!",
                                "Ошибка" , MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }
    }
}