using RANGER.Database;
using System;
using System.Collections.Generic;
using System.Linq;
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

            InitializeTableComboBox();
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
        private void InitializeTableComboBox()
        {
            // Добавляем таблицы в ComboBox
            TableComboBox.Items.Add("Employees");
            TableComboBox.Items.Add("Clients");
            TableComboBox.Items.Add("Firearms");
            TableComboBox.Items.Add("Issues");
            TableComboBox.Items.Add("Warehouse");
            TableComboBox.SelectedIndex = 0;
        }

        private void TableComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TableComboBox.SelectedItem != null)
            {
                LoadData(TableComboBox.SelectedItem.ToString());
            }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            if (TableComboBox.SelectedItem != null)
            {
                LoadData(TableComboBox.SelectedItem.ToString());
            }
        }

        private void LoadData(string tableName)
        {
            try
            {
                switch (tableName)
                {
                    case "Employees":
                        var employees = _database.ExecuteQuery<Employee>("SELECT * FROM Employee");
                        DataGrid.ItemsSource = employees;
                        break;
                    case "Clients":
                        var clients = _database.ExecuteQuery<Client>("SELECT * FROM Client");
                        DataGrid.ItemsSource = clients;
                        break;
                    case "Firearms":
                        var firearms = _database.ExecuteQuery<Firearm>(@"SELECT * FROM Firearm");
                        DataGrid.ItemsSource = firearms;
                        break;
                    case "Issues":
                        var issues = _database.ExecuteQuery<Issue>(@"SELECT * FROM Issue");
                        DataGrid.ItemsSource = issues;
                        break;
                    case "Warehouse":
                        var warehouse = _database.ExecuteQuery<Warehouse>(@"SELECT * FROM Warehouse");
                        DataGrid.ItemsSource = warehouse;
                        break;
                }

                StatusText.Text = $"Данные загружены из таблицы: {tableName}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}", "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "Ошибка при загрузке данных";
            }
        }
    }
}
