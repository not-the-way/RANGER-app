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

        public MainWindow(Employee employee)
        {
            InitializeComponent();

            LoggedUserTextBox.Text = $"Пользователь: {employee.FullName}";
            AccessLevelTextBox.Text = $"Вход как: {employee.AccessLevel}";

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
                    
                    ClientsRBtn.Visibility = Visibility.Visible;
                    FirearmsRBtn.Visibility = Visibility.Visible;
                    IssueRBtn.Visibility = Visibility.Visible;
                    WarehouseRBtn.Visibility = Visibility.Visible;
                    EmployeesRBtn.Visibility = Visibility.Visible;
                    break;
                    

                case "Manager":
                    
                    ClientsRBtn.Visibility = Visibility.Visible;
                    FirearmsRBtn.Visibility = Visibility.Visible;
                    IssueRBtn.Visibility = Visibility.Visible;
                    WarehouseRBtn.Visibility = Visibility.Visible;
                    EmployeesRBtn.Visibility = Visibility.Visible;
                    break;

                case "Warehouse":
                    FirearmsRBtn.Visibility = Visibility.Visible;
                    WarehouseRBtn.Visibility = Visibility.Visible;
                    break;

                case "Issue":
                    IssueRBtn.Visibility = Visibility.Visible;
                    ClientsRBtn.Visibility = Visibility.Visible;
                    break;
            }
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

        }

        private void FirearmsRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Firearm";
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
                    //case "Client":
                    //    LoadDataForClients();
                    //    break;
                    case "Firearm":
                        LoadDataForFirearms();
                        break;
                        //case "Issue":
                        //    LoadDataForIssues();
                        //    break;
                        //case "Warehouse":
                        //    LoadDataForWarehouse();
                        //    break;
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
