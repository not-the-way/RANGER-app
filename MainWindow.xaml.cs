using RANGER.Database;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static RANGER.Database.DatabaseImplementation;

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            UpdateWhoIsLoggedStatus();
        }


        // Метод должен изменять текст в LoggedUserTextBox на имя пользователя
        public void UpdateWhoIsLoggedStatus()
        {
            var whoIsLogged = LoggedUserTextBox;
            var dbNameReader = new DatabaseImplementation().ExecuteQuery<Employee>(@"SELECT * FROM Employee;");

            // string userFullName = dbNameReader.Where(n => n.FullName.Equals("");
            // whoIsLogged.Text = $"Пользователь: {userFullName}";
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }

        // Код ниже отвечает за выход из приложения
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // При закрытии окна показывается MessageBox с кнопками да и нет
            var mBox = MessageBox.Show("Вы точно хотите выйти?",
                                       "Выход из приложения",
                                       MessageBoxButton.YesNo,
                                       MessageBoxImage.Question,
                                       MessageBoxResult.No);

            // Здесь нужно реализовать выход приложения при нажатии "Да"
            // И продолжение работы приложения при "Нет"
            if (mBox == MessageBoxResult.Yes)
            {
                Application.Current.Shutdown();
            }
            else
            {

            }
        }
    }
}
