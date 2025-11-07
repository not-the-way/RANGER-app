using RANGER.Database;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
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
using static RANGER.Database.DataBaseConnection;
using static RANGER.Database.DataBaseModel;

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataBaseConnection _database;

        public MainWindow(Employee employee)
        {
            InitializeComponent();

            LoggedUserTextBox.Text = $"Пользователь: {employee.FullName}";
            AccessLevelTextBox.Text = $"Вход как: {employee.AccessLevel}";

            _database = new DataBaseConnection();
            LoadData();
        }

        // Просто на всякий случай
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

            // Здесь реализуется условие, приложение останавливается при нажатии "Да" и работа продолжается при "Нет"
            if (mBox == MessageBoxResult.Yes)
            {
                Application.Current.Shutdown();
            }
            else
            {
                e.Cancel = true;
            }
        }

        // РАБОТА С БД:

        private void LoadData()
        {
            try
            {
                var employee = _database.ExecuteQuery<Employee>("SELECT * FROM Employee");
                var gridView = (GridView)dataListView.View;

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Номер сотрудника",
                    DisplayMemberBinding = new Binding("Employee_ID")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "ФИО",
                    DisplayMemberBinding = new Binding("FullName")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Дата найма",
                    DisplayMemberBinding = new Binding("EmploymentDate")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Логин",
                    DisplayMemberBinding = new System.Windows.Data.Binding("Login"),
                    Width = 100
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Уровень доступа",
                    DisplayMemberBinding = new System.Windows.Data.Binding("AccessLevel"),
                    Width = 100
                });

                dataListView.ItemsSource = employee;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
