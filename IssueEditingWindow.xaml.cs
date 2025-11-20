using RANGER.Database;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using static RANGER.Database.DataBaseModel;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для IssueEditingWindow.xaml
    /// </summary>
    public partial class IssueEditingWindow : Window
    {
        private Issue issue;
        private DataBaseConnection db;
        private bool isNew;

        private List<Client> clients;
        private List<Employee> employees;
        private List<Firearm> firearms;
        private List<Warehouse> warehouseItems;
        private List<Issue> activeIssues;


        public IssueEditingWindow(Issue issue, DataBaseConnection db)
        {
            InitializeComponent();

            this.issue = issue;
            this.db = db;
            this.isNew = issue == null;

            if (isNew)
            {
                this.issue = new Issue();
                Title = "Добавление выдачи";
            }
            else
            {
                Title = "Редактирование выдачи";
            }

            LoadComboBoxData();
            LoadIssueData();
            //UpdateSelectionInfo();
        }

        private void LoadComboBoxData()
        {
            try
            {
                activeIssues = db.ExecuteQuery<Issue>(
                    "SELECT * FROM Issue WHERE DateTimeOfReturn > @currentDate",
                    new SQLiteParameter[] { new SQLiteParameter("@currentDate", DateTime.Now) });

                // Загрузка клиентов
                clients = db.ExecuteQuery<Client>("SELECT * FROM Client ORDER BY FullName");
                cmbClient.ItemsSource = clients;

                // Загрузка сотрудников
                employees = db.ExecuteQuery<Employee>("SELECT * FROM Employee " +
                                                      "WHERE Employee.AccessLevel = 'Issue' " +
                                                      "ORDER BY FullName");
                cmbEmployee.ItemsSource = employees;

                // Загрузка оружия
                firearms = db.ExecuteQuery<Firearm>(
                    "SELECT f.*, fc.Category, fcond.Condition " +
                    "FROM Firearm f " +
                    "LEFT JOIN FirearmCategory fc ON f.Category = fc.Category " +
                    "LEFT JOIN FirearmCondition fcond ON f.Condition = fcond.Condition " +
                    "WHERE f.Condition = 'Пригодно для использования' " + // Только исправное оружие
                    "ORDER BY f.Name");
                cmbFirearm.ItemsSource = firearms;

                // Загрузка предметов со склада
                warehouseItems = db.ExecuteQuery<Warehouse>(
                    "SELECT w.*, wt.Type as WarehouseType " +
                    "FROM Warehouse w " +
                    "LEFT JOIN TypesForWarehouse wt ON w.Warehouse_Type = wt.Type " +
                    "WHERE w.Quantity > 0 " + // Только предметы в наличии
                    "ORDER BY w.ItemName");
                cmbWarehouse.ItemsSource = warehouseItems;

                var availableClients = FilterAvailableClients(clients, activeIssues);
                var availableEmployees = FilterAvailableEmployees(employees, activeIssues);
                var availableFireaerms = FilterAvailableFirearms(firearms, activeIssues);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<Client> FilterAvailableClients(List<Client> clients, List<Issue> activeIssues)
        {
            var clientsWithActiveIssues = activeIssues.Select(i => i.Client_ID).Distinct().ToList();
            return clients.Where(c => !clientsWithActiveIssues.Contains(c.Client_ID)).ToList();
        }

        private List<Employee> FilterAvailableEmployees(List<Employee> employees, List<Issue> activeIssues)
        {
            var employeesWithActiveIssues = activeIssues.Select(i => i.Employee_ID).Distinct().ToList();
            return employees.Where(e => !employeesWithActiveIssues.Contains(e.Employee_ID)).ToList();
        }

        private List<Firearm> FilterAvailableFirearms(List<Firearm> firearms, List<Issue> activeIssues)
        {
            var firearmsWithActiveIssues = activeIssues.Select(i => i.Firearm_SerialNumber).Distinct().ToList();
            return firearms.Where(e => !firearmsWithActiveIssues.Contains(e.FirearmSerialNumber)).ToList();
        }

        private void LoadIssueData()
        {
            if (!isNew)
            {
                // Установка выбранных значений в ComboBox
                if (issue.Client_ID > 0)
                {
                    var client = clients.FirstOrDefault(c => c.Client_ID == issue.Client_ID);
                    cmbClient.SelectedItem = client;
                }

                if (issue.Employee_ID > 0)
                {
                    var employee = employees.FirstOrDefault(e => e.Employee_ID == issue.Employee_ID);
                    cmbEmployee.SelectedItem = employee;
                }

                if (!string.IsNullOrEmpty(issue.Firearm_SerialNumber))
                {
                    var firearm = firearms.FirstOrDefault(f => f.FirearmSerialNumber == issue.Firearm_SerialNumber);
                    cmbFirearm.SelectedItem = firearm;
                }

                if (issue.Item_ID > 0)
                {
                    var warehouseItem = warehouseItems.FirstOrDefault(w => w.Item_ID == issue.Item_ID);
                    cmbWarehouse.SelectedItem = warehouseItem;
                }

                dpDateTimeOfIssue.SelectedDate = issue.DateTimeOfIssue;
                dpDateTimeOfReturn.SelectedDate = issue.DateTimeOfReturn;
            }
            else
            {
                // Установка значений по умолчанию для новой выдачи
                dpDateTimeOfIssue.SelectedDate = DateTime.Now;
                dpDateTimeOfReturn.SelectedDate = DateTime.Now.AddDays(1);
            }
        }

        //private void UpdateSelectionInfo()
        //{
        //    var info = new StringBuilder();

        //    if (cmbClient.SelectedItem is Client selectedClient)
        //    {
        //        info.AppendLine($"Клиент: {selectedClient.FullName}");
        //        info.AppendLine($"Паспорт: {selectedClient.Passport}");
        //        info.AppendLine($"Совершеннолетний: {(selectedClient.IsMature ? "Да" : "Нет")}");
        //        info.AppendLine($"Отказ от ответственности: {(selectedClient.ReleaseOfLiability ? "Да" : "Нет")}");
        //        info.AppendLine();
        //    }

        //    if (cmbFirearm.SelectedItem is Firearm selectedFirearm)
        //    {
        //        info.AppendLine($"Оружие: {selectedFirearm.Name}");
        //        info.AppendLine($"Серийный номер: {selectedFirearm.FirearmSerialNumber}");
        //        info.AppendLine($"Категория: {selectedFirearm.Category}");
        //        info.AppendLine($"Состояние: {selectedFirearm.Condition}");
        //        info.AppendLine();
        //    }

        //    if (cmbWarehouse.SelectedItem is Warehouse selectedWarehouse)
        //    {
        //        info.AppendLine($"Предмет: {selectedWarehouse.ItemName}");
        //        info.AppendLine($"Тип: {selectedWarehouse.WarehouseType}");
        //        info.AppendLine($"Количество в наличии: {selectedWarehouse.Quantity}");
        //    }

        //    tbSelectionInfo.Text = info.ToString();
        //}

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInput())
                return;

            try
            {
                // Получаем выбранные значения
                var selectedClient = cmbClient.SelectedItem as Client;
                var selectedEmployee = cmbEmployee.SelectedItem as Employee;
                var selectedFirearm = cmbFirearm.SelectedItem as Firearm;
                var selectedWarehouse = cmbWarehouse.SelectedItem as Warehouse;

                if (selectedClient == null || selectedEmployee == null ||
                    selectedFirearm == null || selectedWarehouse == null)
                {
                    MessageBox.Show("Заполните все обязательные поля", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Обновляем данные выдачи
                issue.Client_ID = selectedClient.Client_ID;
                issue.Employee_ID = selectedEmployee.Employee_ID;
                issue.Firearm_SerialNumber = selectedFirearm.FirearmSerialNumber;
                issue.Item_ID = selectedWarehouse.Item_ID;
                issue.DateTimeOfIssue = dpDateTimeOfIssue.SelectedDate.Value;
                issue.DateTimeOfReturn = dpDateTimeOfReturn.SelectedDate.Value;

                if (isNew)
                {
                    // Вставка новой записи
                    db.Query(
                        "INSERT INTO Issue (Client_ID, Employee_ID, Firearm_SerialNumber, Item_ID, DateTimeOfIssue, DateTimeOfReturn) " +
                        "VALUES (@clientId, @employeeId, @firearmSn, @itemId, @issueDate, @returnDate)",
                        new SQLiteParameter[] {
                            new SQLiteParameter("@clientId", issue.Client_ID),
                            new SQLiteParameter("@employeeId", issue.Employee_ID),
                            new SQLiteParameter("@firearmSn", issue.Firearm_SerialNumber),
                            new SQLiteParameter("@itemId", issue.Item_ID),
                            new SQLiteParameter("@issueDate", issue.DateTimeOfIssue),
                            new SQLiteParameter("@returnDate", issue.DateTimeOfReturn)
                        });
                }
                else
                {
                    // Обновление существующей записи
                    db.Query(
                        "UPDATE Issue SET Client_ID = @clientId, Employee_ID = @employeeId, " +
                        "Firearm_SerialNumber = @firearmSn, Item_ID = @itemId, " +
                        "DateTimeOfIssue = @issueDate, DateTimeOfReturn = @returnDate " +
                        "WHERE Issue_ID = @id",
                        new SQLiteParameter[] {
                            new SQLiteParameter("@clientId", issue.Client_ID),
                            new SQLiteParameter("@employeeId", issue.Employee_ID),
                            new SQLiteParameter("@firearmSn", issue.Firearm_SerialNumber),
                            new SQLiteParameter("@itemId", issue.Item_ID),
                            new SQLiteParameter("@issueDate", issue.DateTimeOfIssue),
                            new SQLiteParameter("@returnDate", issue.DateTimeOfReturn),
                            new SQLiteParameter("@id", issue.Issue_ID)
                        });
                }

                // Обновляем количество на складе (уменьшаем на 1)
                if (isNew)
                {
                    db.Query(
                        "UPDATE Warehouse SET Quantity = Quantity - 1 WHERE Item_ID = @itemId",
                        new SQLiteParameter[] { new SQLiteParameter("@itemId", issue.Item_ID) });
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ValidateInput()
        {
            if (cmbClient.SelectedItem == null)
            {
                MessageBox.Show("Выберите клиента", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (cmbEmployee.SelectedItem == null)
            {
                MessageBox.Show("Выберите сотрудника", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (cmbFirearm.SelectedItem == null)
            {
                MessageBox.Show("Выберите оружие", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (cmbWarehouse.SelectedItem == null)
            {
                MessageBox.Show("Выберите предмет со склада", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (dpDateTimeOfIssue.SelectedDate == null)
            {
                MessageBox.Show("Выберите дату выдачи", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (dpDateTimeOfReturn.SelectedDate <= dpDateTimeOfIssue.SelectedDate)
            {
                MessageBox.Show("Дата возврата должна быть позже даты выдачи", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // Проверка возраста клиента
            var client = cmbClient.SelectedItem as Client;
            if (client != null && !client.IsMature)
            {
                var result = MessageBox.Show("Клиент несовершеннолетний. Продолжить сохранение?",
                    "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No)
                    return false;
            }

            return true;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}