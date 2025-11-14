using RANGER.Database;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SQLite;
using System.Globalization;
using System.Linq;
using System.Net.NetworkInformation;
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

            currentTable = null;

            _database = new DataBaseConnection();

            ClientsRBtn.Visibility = Visibility.Collapsed;
            FirearmsRBtn.Visibility = Visibility.Collapsed;
            IssueRBtn.Visibility = Visibility.Collapsed;
            WarehouseRBtn.Visibility = Visibility.Collapsed;
            EmployeesRBtn.Visibility = Visibility.Collapsed;
            btnAdd.Visibility = Visibility.Collapsed;
            btnEdit.Visibility = Visibility.Collapsed;
            btnDelete.Visibility = Visibility.Collapsed;

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
                    FirearmsRBtn.Visibility = Visibility.Visible;
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


            switch (currentAccessLevel)
            {
                case "Сис. Админ":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    btnDelete.Visibility = Visibility.Visible;
                    break;
            }
        }
        private void ClientsRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Client";
            LoadData();

            switch (currentAccessLevel)
            {
                case "Сис. Админ":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    btnDelete.Visibility = Visibility.Visible;
                    break;

                case "Сотрудник выдачи":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void FirearmsRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Firearm";
            LoadData();

            switch (currentAccessLevel)
            {
                case "Сис. Админ":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    btnDelete.Visibility = Visibility.Visible;
                    break;

                case "Сотрудник склада":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void IssueRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Issue";
            LoadData();

            switch (currentAccessLevel)
            {
                case "Сис. Админ":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    btnDelete.Visibility = Visibility.Visible;
                    break;

                case "Сотрудник выдачи":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void WarehouseRBtn_Checked(object sender, RoutedEventArgs e)
        {
            currentTable = "Warehouse";
            LoadData();

            switch (currentAccessLevel)
            {
                case "Сис. Админ":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    btnDelete.Visibility = Visibility.Visible;
                    break;

                case "Сотрудник склада":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                switch (currentTable)
                {
                    case "Employee":
                        var employeeWindow = new EmployeeEditingWindow(null, _database);
                        if (employeeWindow.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;

                    case "Client":
                        var clientWindow = new ClientEditingWindow(null, _database);
                        if (clientWindow.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;

                    case "Firearm":
                        var firearmWindow = new FirearmEditingWindow(null, _database);
                        if (firearmWindow.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;

                    case "Issue":
                        var issueWindow = new IssueEditingWindow(null, _database);
                        if (issueWindow.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;

                    case "Warehouse":
                        var warehouseWindow = new WarehouseEditingWindow(null, _database);
                        if (warehouseWindow.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (dataListView.SelectedItem == null)
            {
                MessageBox.Show("Выберите запись для редактирования", "Информация",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                switch (currentTable)
                {
                    case "Employee":
                        var employee = (Employee)dataListView.SelectedItem;
                        var employeeWindow = new EmployeeEditingWindow(employee, _database);
                        if (employeeWindow.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;

                    case "Client":
                        var client = (Client)dataListView.SelectedItem;
                        var clientWindw = new ClientEditingWindow(client, _database);
                        if (clientWindw.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;

                    case "Firearm":
                        var firearm = (Firearm)dataListView.SelectedItem;
                        var firearmWindow = new FirearmEditingWindow(firearm, _database);
                        if (firearmWindow.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;

                    case "Issue":
                        var issue = (Issue)dataListView.SelectedItem;
                        var issueWindow = new IssueEditingWindow(issue, _database);
                        if (issueWindow.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;

                    case "Warehouse":
                        var warehouse = (Warehouse)dataListView.SelectedItem;
                        var warehouseWindow = new WarehouseEditingWindow(warehouse, _database);
                        if (warehouseWindow.ShowDialog() == true)
                        {
                            LoadData();
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при редактировании: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (dataListView.SelectedItem == null)
            {
                MessageBox.Show("Выберите запись для удаления", "Информация",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var result = MessageBox.Show("Вы уверены, что хотите удалить выбранную запись?",
                "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    switch (currentTable)
                    {
                        case "Employee":
                            var employee = (Employee)dataListView.SelectedItem;
                            _database.Query("DELETE FROM Employee WHERE Employee_ID = @id",
                                new SQLiteParameter[] { new SQLiteParameter("@id", employee.Employee_ID) });
                            break;

                        case "Client":
                            var client = (Client)dataListView.SelectedItem;
                            _database.Query("DELETE FROM Client WHERE Client_ID = @id",
                                new SQLiteParameter[] { new SQLiteParameter("@id", client.Client_ID) });
                            break;

                        case "Firearm":
                            var firearm = (Firearm)dataListView.SelectedItem;
                            _database.Query("DELETE FROM Firearm WHERE FirearmSerialNumber = @sn",
                                new SQLiteParameter[] { new SQLiteParameter("@sn", firearm.FirearmSerialNumber) });
                            break;

                        case "Warehouse":
                            var warehouse = (Warehouse)dataListView.SelectedItem;
                            _database.Query("DELETE FROM Warehouse WHERE Item_ID = @id",
                                new SQLiteParameter[] { new SQLiteParameter("@id", warehouse.Item_ID) });
                            break;

                        case "Issue":
                            var issue = (Issue)dataListView.SelectedItem;
                            _database.Query("DELETE FROM Issue WHERE Issue_ID = @id",
                                new SQLiteParameter[] { new SQLiteParameter("@id", issue.Issue_ID) });
                            break;
                    }

                    LoadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            if (currentTable == null)
            {
                MessageBox.Show("Выберите таблицу!", "Обновление",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }

            LoadData();
        }

        private void dataListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            btnEdit.IsEnabled = dataListView.SelectedItem != null;
            btnDelete.IsEnabled = dataListView.SelectedItem != null;
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
                var firearm = _database.ExecuteQuery<Firearm>
                    ("SELECT f.*, fc.Category, fcond.Condition " +
                     "FROM Firearm f " +
                     "LEFT JOIN FirearmCategory fc ON f.Category = fc.Category " +
                     "LEFT JOIN FirearmCondition fcond ON f.Condition = fcond.Condition " +
                     "ORDER BY f.FirearmSerialNumber"
                    );
                var gridView = (GridView)dataListView.View;
                gridView.Columns.Clear();

                gridView.Columns.Add(CreateTextColumn("Серийный номер", "FirearmSerialNumber", 140));
                gridView.Columns.Add(CreateTextColumn("Название", "Name", 150));
                gridView.Columns.Add(CreateTextColumn("Категория", "Category", 175));
                gridView.Columns.Add(CreateTextColumn("Состояние", "Condition", 180));
                gridView.Columns.Add(CreateDateColumn("Посл. техобслуживание", "LastMaitenanceDate", 165));

                dataListView.ItemContainerStyle = CreateFirearmItemStyle();
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
                gridView.Columns.Clear();

                gridView.Columns.Add(CreateTextColumn("ФИО", "FullName", 200));
                gridView.Columns.Add(CreateTextColumn("Номер телефона", "PhoneNumber"));
                gridView.Columns.Add(CreateDateColumn("Дата найма", "EmploymentDate"));

                if (currentAccessLevel == "Сис. Админ")
                {
                    gridView.Columns.Add(CreateTextColumn("Логин", "Login", 100));
                    gridView.Columns.Add(CreateTextColumn("Пароль", "Password", 100));
                    gridView.Columns.Add(CreateTextColumn("Уровень доступа", "AccessLevel", 100));
                }

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
                gridView.Columns.Clear();

                gridView.Columns.Add(CreateTextColumn("ФИО клиента", "FullName", 200));
                gridView.Columns.Add(CreateTextColumn("Номер телефона", "PhoneNumber", 200));

                if (currentAccessLevel == "Сис. Админ" || currentAccessLevel == "Менеджер")
                {
                    gridView.Columns.Add(CreateTextColumn("Серия и номер паспорта", "Passport", 70));
                }

                gridView.Columns.Add(CreateBoolColumn("Совершеннолетний", "IsMature", 70));
                gridView.Columns.Add(CreateBoolColumn("Отказ от ответственности", "ReleaseOfLiability"));

                dataListView.ItemContainerStyle = CreateClientLiabilityStyle();
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
                ("SELECT i.Issue_ID, e.FullName as Employee_FullName, c.FullName as Client_FullName, " +
                 "f.Name as Firearm_Name, w.ItemName as Warehouse_ItemName, i.DateTimeOfIssue, i.DateTimeOfReturn " +
                 "FROM Issue i " +
                 "LEFT JOIN Client c ON i.Client_ID = c.Client_ID " +
                 "LEFT JOIN Employee e ON i.Employee_ID = e.Employee_ID " +
                 "LEFT JOIN Firearm f ON i.Firearm_SerialNumber = f.FirearmSerialNumber " +
                 "LEFT JOIN Warehouse w ON i.Item_ID = w.Item_ID "
                );

                var gridView = (GridView)dataListView.View;

                gridView.Columns.Add(CreateTextColumn("ФИО сотрудника", "Employee.FullName", 150));
                gridView.Columns.Add(CreateTextColumn("ФИО клиента", "Client.FullName", 150));
                gridView.Columns.Add(CreateTextColumn("Название оружия", "Firearm.Name", 150));
                gridView.Columns.Add(CreateTextColumn("Название предмета", "Warehouse.ItemName", 150));
                gridView.Columns.Add(CreateDateTimeColumn("Дата выдачи", "DateTimeOfIssue"));
                gridView.Columns.Add(CreateDateTimeColumn("Дата возврата", "DateTimeOfReturn"));

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
                 "LEFT JOIN TypesForWarehouse wt ON w.Warehouse_Type = wt.Type " +
                 "ORDER BY w.Item_ID");

                var gridView = (GridView)dataListView.View;

                gridView.Columns.Add(CreateTextColumn("Наименование", "ItemName"));
                gridView.Columns.Add(CreateDateTimeColumn("Дата поставки", "DeliveryDate"));
                gridView.Columns.Add(CreateTextColumn("Тип предмета", "Warehouse_Type"));
                gridView.Columns.Add(CreateTextColumn("Количество", "Quantity"));

                dataListView.ItemsSource = warehouse;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Форматы колонок
        private GridViewColumn CreateDateColumn(string header, string bindingPath, int width = 100)
        {
            return new GridViewColumn
            {
                Header = header,
                Width = width,
                DisplayMemberBinding = new Binding(bindingPath)
                {
                    StringFormat = "dd.MM.yyyy"
                }
            };
        }

        private GridViewColumn CreateDateTimeColumn(string header, string bindingPath, int width = 100)
        {
            return new GridViewColumn
            {
                Header = header,
                Width = width,
                DisplayMemberBinding = new Binding(bindingPath)
                {
                    StringFormat = "dd.MM.yyyy HH:mm"
                }
            };
        }

        private GridViewColumn CreateTextColumn(string header, string bindingPath, int width = 100)
        {
            return new GridViewColumn
            {
                Header = header,
                DisplayMemberBinding = new Binding(bindingPath),
                Width = width
            };
        }

        private GridViewColumn CreateBoolColumn(string header, string bindingPath, int width = 100)
        {
            var dataTemplate = new DataTemplate();
            var factory = new FrameworkElementFactory(typeof(TextBlock));

            // Создаем привязку с конвертером
            var binding = new Binding(bindingPath)
            {
                Converter = new BooleanToTextConverter()
            };

            factory.SetBinding(TextBlock.TextProperty, binding);
            dataTemplate.VisualTree = factory;

            return new GridViewColumn
            {
                Header = header,
                Width = width,
                CellTemplate = dataTemplate
            };
        }

        public class BooleanToTextConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                if (value is bool boolValue)
                {
                    return boolValue ? "Да" : "Нет";
                }
                return "Нет";
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            {
                if (value is string stringValue)
                {
                    return stringValue == "Да";
                }
                return false;
            }
        }

        // Не поверите: ЭТО СТИЛИ!!

        private Style CreateFirearmItemStyle()
        {
            var style = new Style(typeof(ListViewItem));

            // Триггер для состояния "Требуется техобслуживание"
            var trigger = new DataTrigger
            {
                Binding = new Binding("Condition"),
                Value = "Требуется техобслуживание"
            };

            trigger.Setters.Add(new Setter(BackgroundProperty, Brushes.LightYellow));
            trigger.Setters.Add(new Setter(BorderBrushProperty, Brushes.Orange));
            trigger.Setters.Add(new Setter(ToolTipProperty, "Требуется техническое обслуживание"));
            trigger.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(1)));

            style.Triggers.Add(trigger);

            return style;
        }

        private Style CreateClientLiabilityStyle()
        {
            var style = new Style(typeof(ListViewItem));

            // Триггер для проверки значения отказа от ответственности
            var trigger = new DataTrigger
            {
                Binding = new Binding("ReleaseOfLiability"),
                Value = false
            };

            trigger.Setters.Add(new Setter(BackgroundProperty, Brushes.LightYellow));
            trigger.Setters.Add(new Setter(BorderBrushProperty, Brushes.Orange));
            trigger.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(1)));
            trigger.Setters.Add(new Setter(ToolTipProperty, "Необходимо подписать отказ от ответственности"));

            style.Triggers.Add(trigger);

            return style;
        }

        private protected bool EasterEgg(bool value)
        {
            if (!value)
            {
                value = true;
                
            }
            if (true == true)
            {
                return false;
            }
            else
            {
                Application.Current.Shutdown();
            }
        }
    }
}