using RANGER.Database;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        private string currentTable;
        private string currentAccessLevel;

        public MainWindow(Employee employee)
        {
            InitializeComponent();

            LoggedUserTextBox.Text = $"Пользователь: {employee.FullName}";

            currentTable = "Employee";

            _database = new DataBaseConnection();

            ClientsRBtn.Visibility = Visibility.Collapsed;
            FirearmsRBtn.Visibility = Visibility.Collapsed;
            IssueRBtn.Visibility = Visibility.Collapsed;
            WarehouseRBtn.Visibility = Visibility.Collapsed;
            EmployeesRBtn.Visibility = Visibility.Collapsed;

            // Работа с уровнями доступа (WIP)
            switch (employee.AccessLevel)
            {
                case "Admin":
                    currentAccessLevel = "Сис. Админ";

                    ClientsRBtn.Visibility = Visibility.Visible;
                    FirearmsRBtn.Visibility = Visibility.Visible;
                    IssueRBtn.Visibility = Visibility.Visible;
                    WarehouseRBtn.Visibility = Visibility.Visible;
                    EmployeesRBtn.Visibility = Visibility.Visible;
                    break;
                    

                case "Manager":
                    currentAccessLevel = "Менеджер тира";

                    ClientsRBtn.Visibility = Visibility.Visible;
                    FirearmsRBtn.Visibility = Visibility.Visible;
                    IssueRBtn.Visibility = Visibility.Visible;
                    WarehouseRBtn.Visibility = Visibility.Visible;
                    EmployeesRBtn.Visibility = Visibility.Visible;
                    break;

                case "Warehouse":
                    currentAccessLevel = "Сотрудник склада";

                    FirearmsRBtn.Visibility = Visibility.Visible;
                    WarehouseRBtn.Visibility = Visibility.Visible;
                    break;

                case "Issue":
                    currentAccessLevel = "Сотрудник выдачи";

                    IssueRBtn.Visibility = Visibility.Visible;
                    ClientsRBtn.Visibility = Visibility.Visible;
                    break;
            }

            AccessLevelTextBox.Text = $"Вход как: {currentAccessLevel}";
        }

        // Код ниже отвечает за выход из приложения
        private void Window_Closing(object sender, CancelEventArgs e)
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

        // Кнопка выхода (WIP)
        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            LoginWindow loginWindow = new LoginWindow();
            this.Close();
            loginWindow.Show();
        }

        // РАБОТА С БД:

        private void EmployeesRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Employee";
            LoadData();
        }
        private void ClientsRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Client";
            LoadData();
        }

        private void FirearmsRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Firearm";
            LoadData();
        }

        private void IssueRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Issue";
            LoadData();
        }

        private void WarehouseRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Warehouse";
            LoadData();
        }

        private void LoadData()
        {
            try
            {
                dataListView.ItemsSource = null;
                dataListView.View = new GridView();
                var gridView = (GridView)dataListView.View;
                gridView.Columns.Clear();

                switch (currentTable)
                {
                    case "Employee":
                        LoadDataForEmployees();
                        break;
                    case "Client":
                        LoadDataForClients();
                        break;
                    case "Firearm":
                        LoadDataForFirearms();
                        break;
                    case "Issue":
                        LoadDataForIssues();
                        break;
                    case "Warehouse":
                        LoadDataForWarehouse();
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDataForFirearms()
        {
            try
            {
                var firearm = _database.ExecuteQuery<Firearm>("SELECT * FROM Firearm");
                var gridView = (GridView)dataListView.View;

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Серийный номер",
                    DisplayMemberBinding = new Binding("FirearmSerialNumber")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Название",
                    DisplayMemberBinding = new Binding("Name")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Категория",
                    DisplayMemberBinding = new Binding("Category")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Состояние",
                    DisplayMemberBinding = new Binding("Condition")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Дата последнего обслуживания",
                    DisplayMemberBinding = new Binding("LastMaitenanceDate")
                });

                dataListView.ItemsSource = firearm;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDataForEmployees()
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
                    DisplayMemberBinding = new Binding("Login"),
                    Width = 100
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Уровень доступа",
                    DisplayMemberBinding = new Binding("AccessLevel"),
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

        private void LoadDataForClients()
        {
            try
            {
                var clients = _database.ExecuteQuery<Client>("SELECT * FROM Client");
                var gridView = (GridView)dataListView.View;

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Номер клиента",
                    DisplayMemberBinding = new Binding("Client_ID")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "ФИО клиента",
                    DisplayMemberBinding = new Binding("FullName")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Серия и номер паспорта",
                    DisplayMemberBinding = new Binding("Passport")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Совершеннолетний",
                    DisplayMemberBinding = new Binding("IsMature")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Подписан отказ от ответственности",
                    DisplayMemberBinding = new Binding("ReleaseOfLiability")
                });

                dataListView.ItemsSource = clients;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDataForIssues()
        {
            try
            {
                var issues = _database.ExecuteQuery<Issue>
                ("SELECT i.Issue_ID, e.FullName as Employee_FullName, c.Fullname as Client_Fullname, " +
                "f.Name as Firearm_Name, w.ItemName, i.DateTimeOfIssue, i.DateTimeOfReturn " +
                "FROM Issue i " +
                "LEFT JOIN Client c ON i.Client_ID = c.Client_ID " +
                "LEFT JOIN Employee e ON i.Employee_ID = e.Employee_ID " +
                "LEFT JOIN Firearm f ON i.Firearm_SerialNumber = f.FirearmSerialNumber " +
                "LEFT JOIN Warehouse w ON i.Item_ID = w.Item_ID ");
                
                var gridView = (GridView)dataListView.View;

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Номер выдачи",
                    DisplayMemberBinding = new Binding("Issue_ID")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "ФИО сотрудника",
                    DisplayMemberBinding = new Binding("Employee.FullName")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "ФИО клиента",
                    DisplayMemberBinding = new Binding("Client.FullName")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Название оружия",
                    DisplayMemberBinding = new Binding("Firearm.Name")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Название предмета",
                    DisplayMemberBinding = new Binding("Warehouse.ItemName")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Дата выдачи",
                    DisplayMemberBinding = new Binding("DateTimeOfIssue")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Дата возврата",
                    DisplayMemberBinding = new Binding("DateTimeOfReturn")
                });

                dataListView.ItemsSource = issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDataForWarehouse()
        {
            try
            {
                var warehouse = _database.ExecuteQuery<Warehouse>
                ("SELECT w.*, wt.Type as WarehouseType " +
                "FROM Warehouse w " +
                "LEFT JOIN WarehouseType wt ON w.WarehouseType = wt.Type " +
                "ORDER BY w.Item_ID");

                var gridView = (GridView)dataListView.View;

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Номер предмета",
                    DisplayMemberBinding = new Binding("Item_ID")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Наименование",
                    DisplayMemberBinding = new Binding("ItemName")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Дата поставки",
                    DisplayMemberBinding = new Binding("DeliveryDate")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Тип предмета",
                    DisplayMemberBinding = new Binding("WarehouseType.Type")
                });

                gridView.Columns.Add(new GridViewColumn
                {
                    Header = "Количество",
                    DisplayMemberBinding = new Binding("Quantity")
                });

                dataListView.ItemsSource = warehouse;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}