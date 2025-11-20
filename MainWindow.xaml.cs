using RANGER.Database;
using System;
using System.ComponentModel;
using System.Data.SQLite;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using static RANGER.Database.DataBaseModel;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
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
            btnExportToWord.Visibility = Visibility.Hidden;
            btnExportToExcel.Visibility = Visibility.Hidden;

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
                    WarehouseRBtn.Visibility = Visibility.Visible;
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
            if (mBox == MessageBoxResult.Yes)
            {
                System.Windows.Application.Current.Shutdown();
            }
            else
            {
                e.Cancel = true;
            }
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
            btnExportToWord.Visibility = Visibility.Collapsed;
            btnExportToExcel.Visibility = Visibility.Visible;
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
                case "Менеджер тира":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    break;
            }
            btnExportToWord.Visibility = Visibility.Collapsed;
            btnExportToExcel.Visibility = Visibility.Visible;
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
                    btnExportToExcel.Visibility= Visibility.Visible;
                    break;

                case "Сотрудник склада":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    btnExportToExcel.Visibility = Visibility.Visible;
                    break;

                case "Менеджер тира":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility= Visibility.Visible;
                    btnExportToExcel.Visibility = Visibility.Visible;
                    break;
            }
            btnExportToWord.Visibility = Visibility.Collapsed;
            btnExportToExcel.Visibility = Visibility.Visible;
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
                    btnExportToWord.Visibility = Visibility.Visible;
                    btnExportToExcel.Visibility = Visibility.Visible;
                    break;
                case "Сотрудник выдачи":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    btnExportToWord.Visibility = Visibility.Visible;
                    btnExportToExcel.Visibility = Visibility.Visible;
                    break;
                case "Менеджер тира":
                    btnExportToWord.Visibility = Visibility.Visible;
                    btnExportToExcel.Visibility = Visibility.Visible;
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
                    btnExportToExcel.Visibility = Visibility.Visible;
                    break;

                case "Сотрудник склада":
                    btnAdd.Visibility = Visibility.Visible;
                    btnEdit.Visibility = Visibility.Visible;
                    btnExportToExcel.Visibility = Visibility.Visible;
                    break;

                case "Менеджер тира":
                    btnEdit.Visibility = Visibility.Visible;
                    btnAdd.Visibility = Visibility.Visible;
                    btnExportToExcel.Visibility = Visibility.Visible;
                    break;
            }
            btnExportToWord.Visibility = Visibility.Collapsed;
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
                gridView.Columns.Add(CreateTextColumn("Серийный номер", "FirearmSerialNumber", 150));
                gridView.Columns.Add(CreateTextColumn("Название", "Name", 200));
                gridView.Columns.Add(CreateTextColumn("Категория", "Category", 200));
                gridView.Columns.Add(CreateTextColumn("Состояние", "Condition", 200));
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
                gridView.Columns.Add(CreateTextColumn("Номер телефона", "PhoneNumber", 220));
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
                gridView.Columns.Add(CreateTextColumn("ФИО клиента", "FullName", 250));
                gridView.Columns.Add(CreateTextColumn("Номер телефона", "PhoneNumber", 220));
                if (currentAccessLevel == "Сис. Админ" || currentAccessLevel == "Менеджер")
                {
                    gridView.Columns.Add(CreateTextColumn("Серия и номер паспорта", "Passport", 100));
                }
                gridView.Columns.Add(CreateBoolColumn("Совершеннолетний", "IsMature", 100));
                gridView.Columns.Add(CreateBoolColumn("Отказ от ответст.", "ReleaseOfLiability", 100));
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
                gridView.Columns.Add(CreateTextColumn("ФИО сотрудника", "Employee.FullName", 210));
                gridView.Columns.Add(CreateTextColumn("ФИО клиента", "Client.FullName", 210));
                gridView.Columns.Add(CreateTextColumn("Название оружия", "Firearm.Name", 190));
                gridView.Columns.Add(CreateTextColumn("Название предмета", "Warehouse.ItemName", 170));
                gridView.Columns.Add(CreateDateTimeColumn("Дата выдачи", "DateTimeOfIssue", 125));
                gridView.Columns.Add(CreateDateTimeColumn("Дата возврата", "DateTimeOfReturn", 125));
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
                gridView.Columns.Add(CreateTextColumn("Наименование", "ItemName", 250));
                gridView.Columns.Add(CreateDateTimeColumn("Дата поставки", "DeliveryDate", 125));
                gridView.Columns.Add(CreateTextColumn("Тип предмета", "Warehouse_Type", 125));
                gridView.Columns.Add(CreateTextColumn("Количество", "Quantity"));

                dataListView.ItemsSource = warehouse;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void btnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dataListView.Items.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта", "Информация",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                excelApp.ScreenUpdating = true;

                workbook = excelApp.Workbooks.Add();
                worksheet = workbook.ActiveSheet;
                worksheet.Name = GetExcelWorksheetName(currentTable);
                var gridView = dataListView.View as GridView;
                int columnCount = gridView.Columns.Count;
                // Создаем заголовки
                for (int i = 0; i < columnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = gridView.Columns[i].Header.ToString();
                    // Форматирование заголовков
                    var headerCell = worksheet.Cells[1, i + 1];
                    headerCell.Font.Bold = true;
                    headerCell.Font.Size = 12;
                    headerCell.Interior.Color = Excel.XlRgbColor.rgbLightSteelBlue;
                    headerCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    headerCell.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    headerCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    headerCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
                // Заполняем данные с учетом типов
                int row = 2;
                foreach (var item in dataListView.Items)
                {
                    for (int col = 0; col < columnCount; col++)
                    {
                        var column = gridView.Columns[col];
                        object cellValue = GetFormattedCellValue(item, column);

                        var cell = worksheet.Cells[row, col + 1];
                        cell.Value = cellValue;
                        cell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        cell.Borders.Weight = Excel.XlBorderWeight.xlThin;

                        // Форматирование для дат
                        if (cellValue is DateTime)
                        {
                            cell.NumberFormat = "dd.MM.yyyy";
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        // Форматирование для числовых значений
                        if (cellValue is int || cellValue is decimal)
                        {
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        }
                    }
                    row++;
                }
                // Автоподбор ширины столбцов
                worksheet.Columns.AutoFit();
                // Добавляем автофильтр
                if (dataListView.Items.Count > 0)
                {
                    Excel.Range dataRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row - 1, columnCount]];
                    dataRange.AutoFilter(1);
                }
                // Замораживаем область заголовков
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                // Сохраняем файл
                string fileName = $"{GetExcelFileName(currentTable)}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                string filePath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    fileName);
                workbook.SaveAs(filePath);
                MessageBox.Show($"Данные экспортированы в файл: {filePath}", "Экспорт завершен",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Метод для получения имени листа в Excel
        private string GetExcelWorksheetName(string tableName)
        {
            switch (tableName)
            {
                case "Employee":
                    return "Сотрудники";
                case "Client":
                    return "Клиенты";
                case "Firearm":
                    return "Оружие";
                case "Issue":
                    return "Выдачи";
                case "Warehouse":
                    return "Склад";
                default:
                    return "Данные";
            }
        }
        // Метод для получения имени файла
        private string GetExcelFileName(string tableName)
        {
            switch (tableName)
            {
                case "Employee":
                    return "Сотрудники";
                case "Client":
                    return "Клиенты";
                case "Firearm":
                    return "Оружие";
                case "Issue":
                    return "Выдачи";
                case "Warehouse":
                    return "Склад";
                default:
                    return "Данные";
            }
        }
        private object GetFormattedCellValue(object item, GridViewColumn column)
        {
            try
            {
                if (column.DisplayMemberBinding is System.Windows.Data.Binding binding)
                {
                    string propertyPath = binding.Path.Path;
                    string[] properties = propertyPath.Split('.');
                    object currentValue = item;

                    foreach (string prop in properties)
                    {
                        if (currentValue == null) break;

                        var propertyInfo = currentValue.GetType().GetProperty(prop);
                        if (propertyInfo != null)
                        {
                            currentValue = propertyInfo.GetValue(currentValue);
                        }
                        else
                        {
                            currentValue = null;
                            break;
                        }
                    }
                    if (currentValue is DateTime dateValue)
                    {
                        return dateValue.ToString("dd.MM.yyyy");
                    }
                    return currentValue ?? "";
                }
                // Специальная обработка для определенных колонок
                string header = column.Header.ToString();
                switch (header)
                {
                    case "Состояние":
                        if (item is Firearm firearm)
                            return firearm.Condition;
                        break;
                    case "Совершеннолетний":
                        if (item is Client client)
                            return client.IsMature ? "Да" : "Нет";
                        break;
                    case "Отказ от ответст.":
                        if (item is Client client2)
                            return client2.ReleaseOfLiability ? "Да" : "Нет";
                        break;
                    // Обработка дат для конкретных колонок
                    case "Дата найма":
                        if (item is Employee employee)
                            return employee.EmploymentDate;
                        break;
                    case "Последнее обслуж.":
                        if (item is Firearm firearm1)
                            return firearm1.LastMaitenanceDate;
                        break;
                    case "Дата выдачи":
                        if (item is Issue issue1)
                            return issue1.DateTimeOfIssue;
                        break;
                    case "Дата возврата":
                        if (item is Issue issue2)
                            return issue2.DateTimeOfReturn;
                        break;
                    case "Дата поставки":
                        if (item is Warehouse warehouse)
                            return warehouse.DeliveryDate;
                        break;
                    case "DeliveryDate":
                    case "EmploymentDate":
                    case "LastMaitenanceDate":
                    case "DateTimeOfIssue":
                    case "DateTimeOfReturn":
                        if (item != null)
                        {
                            var propertyInfo = item.GetType().GetProperty(header.Replace(" ", ""));
                            if (propertyInfo != null)
                            {
                                var value = propertyInfo.GetValue(item);
                                if (value is DateTime date)
                                {
                                    return date.ToString("dd.MM.yyyy");
                                }
                            }
                        }
                        break;
                }
                return "";
            }
            catch
            { return ""; }
        }
        // Вспомогательный метод для получения значения ячейки
        private string GetCellValue(object item, GridViewColumn column)
        {
            try
            {
                if (column.DisplayMemberBinding is System.Windows.Data.Binding binding)
                {
                    string propertyPath = binding.Path.Path;
                    // Разделяем путь свойства (для сложных путей типа Client.Fullname)
                    string[] properties = propertyPath.Split('.');
                    object currentValue = item;
                    foreach (string prop in properties)
                    {
                        if (currentValue == null) break;
                        var propertyInfo = currentValue.GetType().GetProperty(prop);
                        if (propertyInfo != null)
                        {
                            currentValue = propertyInfo.GetValue(currentValue);
                        }
                        else
                        {
                            currentValue = null;
                            break;
                        }
                    }
                    return currentValue?.ToString() ?? "";
                }
                // Для колонок с CellTemplate (например, Condition с форматированием)
                string header = column.Header.ToString();
                switch (header)
                {
                    case "Состояние":
                        if (item is Firearm firearm)
                            return firearm.Condition;
                        break;
                }
                return "";
            }
            catch
            { return ""; }
        }
        // Кнопка для экспорта в Word
        private void btnExportToWord_Click(object sender, RoutedEventArgs e)
        {
            if (dataListView.SelectedItem is Issue selectedIssue)
            {
                try
                {
                    // По какой-то причине, серийный номер не передается напрямую в колонку
                    string actualSerialNumber = selectedIssue.Firearm_SerialNumber;

                    // Вроде проблему решает
                    if (string.IsNullOrEmpty(actualSerialNumber))
                    {
                        var serialFromDb = _database.ExecuteQuery<Issue>(
                            "SELECT Firearm_SerialNumber FROM Issue WHERE Issue_ID = @issueId",
                            new SQLiteParameter[] { new SQLiteParameter("@issueId", selectedIssue.Issue_ID) })
                            .FirstOrDefault();

                        actualSerialNumber = serialFromDb.Firearm_SerialNumber ?? "Не указан";
                    }

                    // Создание экземпляра Word и нового документа
                    var wordApp = new Word.Application();
                    wordApp.Visible = true;
                    Word._Document wordDoc = wordApp.Documents.Add();

                    // Установка полей страницы
                    wordDoc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(3);
                    wordDoc.PageSetup.RightMargin = wordApp.CentimetersToPoints(3);

                    // Заголовок документа
                    Word.Paragraph titleParagraph = wordDoc.Content.Paragraphs.Add();
                    titleParagraph.Range.Text = "АКТ ВЫДАЧИ ИМУЩЕСТВА СТРЕЛКОВОГО ТИРА";
                    titleParagraph.Range.Font.Bold = 1;
                    titleParagraph.Range.Font.Size = 16;
                    titleParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    titleParagraph.Format.SpaceAfter = 18;
                    titleParagraph.Range.InsertParagraphAfter();

                    // Основной текст согласия
                    Word.Paragraph consentParagraph = wordDoc.Content.Paragraphs.Add();
                    consentParagraph.Range.Text = "Я, нижеподписавшийся ______________________________, " +
                                                 "даю в письменном виде согласие на обработку моих персональных данных, " +
                                                 "и подтверждаю получение инвентаря тира сотрудником данного тира, " +
                                                 "также обязуюсь соблюдать правила, установленные тиром. " +
                                                 "Указанный инвентарь тира:";
                    consentParagraph.Range.Font.Size = 12;
                    consentParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    consentParagraph.Range.Font.Bold = 0;
                    consentParagraph.Format.SpaceAfter = 12;
                    consentParagraph.Range.InsertParagraphAfter();

                    // Создание таблицы с инвентарем
                    Word.Range tableRange = wordDoc.Content;
                    tableRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    Word.Table inventoryTable = wordDoc.Tables.Add(tableRange, 2, 3); // 2 строки (заголовок + данные), 3 колонки

                    // Заголовки таблицы
                    inventoryTable.Cell(1, 1).Range.Text = "Оружие";
                    inventoryTable.Cell(1, 2).Range.Text = "Серийный номер оружия";
                    inventoryTable.Cell(1, 3).Range.Text = "Выданный предмет";

                    // Данные таблицы
                    inventoryTable.Cell(2, 1).Range.Text = selectedIssue.Firearm?.Name ?? "Не указано";
                    inventoryTable.Cell(2, 2).Range.Text = actualSerialNumber;
                    inventoryTable.Cell(2, 3).Range.Text = selectedIssue.Warehouse?.ItemName ?? "Не указан";

                    // Форматирование таблицы
                    inventoryTable.Borders.Enable = 1;
                    inventoryTable.Range.Font.Size = 12;
                    inventoryTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    // Жирный шрифт для заголовков таблицы
                    inventoryTable.Rows[1].Range.Font.Bold = 1;

                    // Автоподбор ширины столбцов под содержимое
                    inventoryTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

                    // Отступ после таблицы
                    inventoryTable.Range.ParagraphFormat.SpaceAfter = 18;

                    // Информация о сотруднике и клиенте
                    Word.Paragraph employeeParagraph = wordDoc.Content.Paragraphs.Add();
                    employeeParagraph.Range.Text = $"Сотрудник выдачи: {selectedIssue.Employee?.FullName ?? "Не указан"}";
                    employeeParagraph.Range.Font.Size = 12;
                    employeeParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    employeeParagraph.Format.SpaceAfter = 6;
                    employeeParagraph.Range.InsertParagraphAfter();

                    Word.Paragraph clientParagraph = wordDoc.Content.Paragraphs.Add();
                    clientParagraph.Range.Text = $"ФИО клиента: {selectedIssue.Client?.FullName ?? "Не указан"}";
                    clientParagraph.Range.Font.Size = 12;
                    clientParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    clientParagraph.Format.SpaceAfter = 24;
                    clientParagraph.Range.InsertParagraphAfter();

                    // Создание таблицы для нижней части с датами и подписями
                    Word.Range footerRange = wordDoc.Content;
                    footerRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    Word.Table footerTable = wordDoc.Tables.Add(footerRange, 3, 2); // 3 строки, 2 колонки

                    // Левая колонка - даты
                    footerTable.Cell(1, 1).Range.Text = $"Дата выдачи: {selectedIssue.DateTimeOfIssue:dd.MM.yyyy HH:mm}";
                    footerTable.Cell(2, 1).Range.Text = $"Дата возврата: {selectedIssue.DateTimeOfReturn:dd.MM.yyyy HH:mm}";

                    // Правая колонка - подписи
                    footerTable.Cell(1, 2).Range.Text = "Подпись клиента: ________________________";
                    footerTable.Cell(2, 2).Range.Text = "Подпись сотрудника: ________________________";

                    // Форматирование нижней таблицы
                    footerTable.Borders.Enable = 0; // Убирает границы таблицы
                    footerTable.Range.Font.Size = 12;

                    // Выравнивание: левая колонка - по левому краю, правая - по правому (ну почти)
                    footerTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    footerTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    footerTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    footerTable.Cell(2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                    // Автоподбор ширины столбцов
                    footerTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

                    // Уборка отступов между ячейками
                    footerTable.Spacing = 0;
                    footerTable.Range.ParagraphFormat.SpaceAfter = 0;
                    footerTable.Range.ParagraphFormat.SpaceBefore = 0;

                    // Форматирование всего документа
                    //FormatDocument(wordDoc);

                    // Сохранение документа с нужным именем
                    string fileName = $"АКТ_ВЫДАЧИ_{selectedIssue.Issue_ID}_{selectedIssue.DateTimeOfIssue:yyyyMMdd}";
                    string filePath = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        fileName + ".docx");

                    wordDoc.SaveAs2(filePath);

                    // Закрытие
                    wordDoc.Close();
                    wordApp.Quit();

                    MessageBox.Show($"Документ сохранен: {filePath}", "Экспорт завершен",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Выберите запись о выдаче для экспорта.", "Информация",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        // Вспомогательный метод для форматированного текста
        private void InsertFormattedText(Word._Document doc, string label, string value)
        {
            if (string.IsNullOrEmpty(value)) return;

            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Range.Text = $"{label}: {value}";
            paragraph.Format.SpaceAfter = 6;
            paragraph.Format.SpaceBefore = 6;
            paragraph.Range.InsertParagraphAfter();
        }
        // Метод для добавления секции с подписями
        private void AddSignaturesSection(Word._Document doc)
        {
            // Добавляем отступ перед подписями
            Word.Paragraph spaceParagraph = doc.Content.Paragraphs.Add();
            spaceParagraph.Format.SpaceAfter = 24;
            spaceParagraph.Range.InsertParagraphAfter();
            // Подпись сотрудника
            Word.Paragraph employeeSignature = doc.Content.Paragraphs.Add();
            employeeSignature.Range.Text = "Сотрудник: _________________________";
            employeeSignature.Format.SpaceAfter = 12;
            employeeSignature.Range.InsertParagraphAfter();
            // Подпись клиента
            Word.Paragraph clientSignature = doc.Content.Paragraphs.Add();
            clientSignature.Range.Text = "Клиент: _________________________";
            clientSignature.Format.SpaceAfter = 12;
            clientSignature.Range.InsertParagraphAfter();
            // Дата получения
            Word.Paragraph dateParagraph = doc.Content.Paragraphs.Add();
            dateParagraph.Range.Text = "Дата получения: _________________________";
            dateParagraph.Format.SpaceAfter = 12;
            dateParagraph.Range.InsertParagraphAfter();
        }
        // Реализация сортировки
        private void dataListView_GridViewColumnHeaderClick(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is GridViewColumnHeader header && header.Column != null)
            {
                string columnName = header.Content.ToString();

                // Получаем компаратор для текущей таблицы
                var customComparer = CustomComparerFactory.GetComparer(currentTable);

                if (customComparer != null && IsCustomSortColumn(currentTable, columnName))
                {
                    ApplyCustomSort(customComparer);
                }
                else
                {
                    // Стандартная сортировка по свойству
                    ApplyStandardSort(GetPropertyNameFromColumn(columnName));
                }
            }
        }
        private bool IsCustomSortColumn(string tableName, string columnName)
        {
            // Определяем, для каких колонок применяем кастомную сортировку
            if (tableName == "Firearm" && columnName == "Состояние")
                return true;
            else if (tableName == "Client" && columnName == "Совершеннолетний")
                return true;
            else if (tableName == "Employee" && columnName == "Уровень доступа")
                return true;
            else if (tableName == "Issue" && columnName == "Дата возврата")
                return true;
            else if (tableName == "Warehouse" && columnName == "Количество")
                return true;
            else
                return false;
        }
        private void ApplyCustomSort(ICustomComparer comparer)
        {
            var listCollectionView = CollectionViewSource.GetDefaultView(dataListView.ItemsSource) as ListCollectionView;

            if (listCollectionView != null && listCollectionView.CanSort)
            {
                listCollectionView.CustomSort = comparer;
            }
        }
        private void ApplyStandardSort(string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName)) return;

            var listCollectionView = CollectionViewSource.GetDefaultView(dataListView.ItemsSource) as ListCollectionView;

            if (listCollectionView != null && listCollectionView.CanSort)
            {
                listCollectionView.CustomSort = null;
                listCollectionView.SortDescriptions.Clear();
                listCollectionView.SortDescriptions.Add(
                    new System.ComponentModel.SortDescription(propertyName,
                    System.ComponentModel.ListSortDirection.Ascending));
            }
        }
        private string GetPropertyNameFromColumn(string columnName)
        {
            switch (columnName)
            {
                case "ФИО клиента":
                    return "FullName";
                case "ФИО сотрудника":
                    return "FullName";
                case "Название":
                    return "ItemName";
                case "Название оружия":
                    return "Name";
                case "Категория":
                    return "Category";
                case "Серийный номер":
                    return "FirearmSerialNumber";
                case "Посл. техобслуживание":
                    return "LastMaitenanceDate";
                case "Дата выдачи":
                    return "DateTimeOfIssue";
                case "Дата возврата":
                    return "DateTimeOfReturn";
                case "Тип предмета":
                    return "Warehouse_Type";
                case "Паспорт":
                    return "Passport";
                case "Дата найма":
                    return "EmploymentDate";
                case "Уровень доступа":
                    return "AccessLevel";
                case "Отказ от ответст.":
                    return "ReleaseOfLiability";
                default:
                    return null;
            }
        }
        public interface ICustomComparer : System.Collections.IComparer
        {
            bool CanSort(Type itemType);
        }
        public static class CustomComparerFactory
        {
            public static ICustomComparer GetComparer(string tableName)
            {
                switch (tableName)
                {
                    case "Firearm":
                        return new FirearmConditionComparer();
                    case "Client":
                        return new ClientComparer();
                    case "Employee":
                        return new EmployeeComparer();
                    case "Issue":
                        return new IssueComparer();
                    case "Warehouse":
                        return new WarehouseComparer();
                    default:
                        return null;
                }
            }
        }
        // Для Firearm (сортировка по состоянию)
        public class FirearmConditionComparer : ICustomComparer
        {
            public bool CanSort(Type itemType) => itemType == typeof(Firearm);

            public int Compare(object x, object y)
            {
                var item1 = x as Firearm;
                var item2 = y as Firearm;

                if (item1 == null || item2 == null) return 0;

                int GetPriority(string condition)
                {
                    switch (condition)
                    {
                        case "Требуется техобслуживание":
                            return 1;
                        case "Пригодно для использования":
                            return 2;
                        default:
                            return 3;
                    }
                }

                return GetPriority(item1.Condition).CompareTo(GetPriority(item2.Condition));
            }
        }
        // Для Client (сортировка по совершеннолетию и ФИО)
        public class ClientComparer : ICustomComparer
        {
            public bool CanSort(Type itemType) => itemType == typeof(Client);

            public int Compare(object x, object y)
            {
                var item1 = x as Client;
                var item2 = y as Client;

                if (item1 == null || item2 == null) return 0;

                // Сначала совершеннолетние, потом несовершеннолетние
                int adultComparison = item2.IsMature.CompareTo(item1.IsMature);
                if (adultComparison != 0) return adultComparison;

                // Затем по ФИО
                return string.Compare(item1.FullName, item2.FullName, StringComparison.OrdinalIgnoreCase);
            }
        }
        // Для Employee (сортировка по уровню доступа)
        public class EmployeeComparer : ICustomComparer
        {
            public bool CanSort(Type itemType) => itemType == typeof(Employee);

            public int Compare(object x, object y)
            {
                var item1 = x as Employee;
                var item2 = y as Employee;

                if (item1 == null || item2 == null) return 0;

                int GetAccessLevelPriority(string level)
                {
                    if (level == null) return 4;

                    switch (level.ToLower())
                    {
                        case "admin":
                            return 1;
                        case "manager":
                            return 2;
                        case "issue":
                            return 3;
                        case "warehouse":
                            return 4;
                        default:
                            return 5;
                    }
                }

                // Сортировка по уровню доступа
                int levelComparison = GetAccessLevelPriority(item1.AccessLevel)
                    .CompareTo(GetAccessLevelPriority(item2.AccessLevel));
                if (levelComparison != 0) return levelComparison;

                // Затем по ФИО
                return string.Compare(item1.FullName, item2.FullName, StringComparison.OrdinalIgnoreCase);
            }
        }
        // Для Issue (сортировка по дате выдачи и статусу)
        public class IssueComparer : ICustomComparer
        {
            public bool CanSort(Type itemType) => itemType == typeof(Issue);

            public int Compare(object x, object y)
            {
                var item1 = x as Issue;
                var item2 = y as Issue;

                if (item1 == null || item2 == null) return 0;

                // Сначала просроченные выдачи
                bool isOverdue1 = item1.DateTimeOfReturn < DateTime.Now;
                bool isOverdue2 = item2.DateTimeOfReturn < DateTime.Now;

                int overdueComparison = isOverdue2.CompareTo(isOverdue1);
                if (overdueComparison != 0) return overdueComparison;

                // Затем по дате выдачи (сначала новые)
                return item2.DateTimeOfIssue.CompareTo(item1.DateTimeOfIssue);
            }
        }
        // Для Warehouse (сортировка по количеству и типу)
        public class WarehouseComparer : ICustomComparer
        {
            public bool CanSort(Type itemType) => itemType == typeof(Warehouse);

            public int Compare(object x, object y)
            {
                var item1 = x as Warehouse;
                var item2 = y as Warehouse;

                if (item1 == null || item2 == null) return 0;

                // Сначала товары с малым количеством
                int quantityComparison = item1.Quantity.CompareTo(item2.Quantity);
                if (quantityComparison != 0) return quantityComparison;

                // Затем по названию
                return string.Compare(item1.ItemName, item2.ItemName, StringComparison.OrdinalIgnoreCase);
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
        // Конвертер для булевых значений из БД
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
        // Стили для выделения
        private System.Windows.Style CreateFirearmItemStyle()
        {
            var style = new System.Windows.Style(typeof(ListViewItem));
            // Триггер для состояния "Требуется техобслуживание"
            var trigger = new DataTrigger
            {
                Binding = new Binding("Condition"),
                Value = "Требуется техобслуживание"
            };
            trigger.Setters.Add(new Setter(BackgroundProperty, Brushes.LightYellow));
            trigger.Setters.Add(new Setter(BorderBrushProperty, Brushes.Red));
            trigger.Setters.Add(new Setter(ToolTipProperty, "Требуется техническое обслуживание"));
            trigger.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(1)));
            style.Triggers.Add(trigger);

            return style;
        }
        private System.Windows.Style CreateClientLiabilityStyle()
        {
            var style = new System.Windows.Style(typeof(ListViewItem));
            // Триггер для проверки значения отказа от ответственности
            var trigger = new DataTrigger
            {
                Binding = new Binding("ReleaseOfLiability"),
                Value = false
            };
            trigger.Setters.Add(new Setter(BackgroundProperty, Brushes.LightYellow));
            trigger.Setters.Add(new Setter(BorderBrushProperty, Brushes.Red));
            trigger.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(1)));
            trigger.Setters.Add(new Setter(ToolTipProperty, "Необходимо подписать отказ от ответственности"));
            style.Triggers.Add(trigger);
            return style;
        }
    }
}