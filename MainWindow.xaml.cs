using RANGER.Database;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SQLite;
using System.Linq;
using System.Reflection;
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
        private string _currentTable;
        private ObservableCollection<object> _currentData;
        private List<object> _originalData;
        private Type _currentType;

        public MainWindow(Employee employee)
        {
            InitializeComponent();

            LoggedUserTextBox.Text = $"Пользователь: {employee.FullName}";
            AccessLevelTextBox.Text = $"Вход как: {employee.AccessLevel}";

            _database = new DataBaseConnection();

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

        private void LoadData(string tableName)
        {
            try
            {
                _currentTable = tableName;
                _currentType = GetTableType(tableName);

                var data = _database.ExecuteQuery($"SELECT * FROM {tableName}", null, _currentType);
                _currentData = new ObservableCollection<object>(data);
                _originalData = new List<object>(data);

                DataGrid.ItemsSource = _currentData;
                StatusText.Text = $"Загружено {_currentData.Count} записей из таблицы: {tableName}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "Ошибка при загрузке данных";
            }
        }

        private Type GetTableType(string tableName)
        {
            switch (tableName)
            {
                case "Employee":
                    return typeof(Employee);
                case "Client":
                    return typeof(Client);
                case "Firearm":
                    return typeof(Firearm);
                case "Issue":
                    return typeof(Issue);
                case "Warehouse":
                    return typeof(Warehouse);
                default:
                    throw new ArgumentException($"Неизвестная таблица: {tableName}");
            }
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var newItem = Activator.CreateInstance(_currentType);
                _currentData.Add(newItem);

                StatusText.Text = "Добавлена новая запись. Не забудьте сохранить изменения!";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении записи: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (DataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите запись для удаления", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var result = MessageBox.Show("Вы уверены, что хотите удалить выбранную запись?",
                                           "Подтверждение удаления",
                                           MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    var selectedItem = DataGrid.SelectedItem;
                    _currentData.Remove(selectedItem);
                    StatusText.Text = "Запись удалена. Не забудьте сохранить изменения!";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении записи: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveChanges();
                StatusText.Text = "Изменения успешно сохранены!";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении изменений: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "Ошибка при сохранении изменений";
            }
        }

        private void SaveChanges()
        {
            // Удаленные записи
            foreach (var originalItem in _originalData)
            {
                if (!_currentData.Contains(originalItem))
                {
                    DeleteRecord(originalItem);
                }
            }

            // Новые и измененные записи
            foreach (var currentItem in _currentData)
            {
                if (!_originalData.Contains(currentItem))
                {
                    InsertRecord(currentItem);
                }
                else
                {
                    UpdateRecord(currentItem);
                }
            }

            // Обновляем оригинальные данные
            _originalData = new List<object>(_currentData);
        }

        private void InsertRecord(object record)
        {
            var properties = GetWritableProperties(record);
            var columnNames = string.Join(", ", properties.Select(p => p.Name));
            var parameterNames = string.Join(", ", properties.Select(p => $"@{p.Name}"));

            var sql = $"INSERT INTO {_currentTable} ({columnNames}) VALUES ({parameterNames})";
            var parameters = CreateParameters(properties, record);

            _database.Query(sql, parameters);
        }

        private void UpdateRecord(object record)
        {
            var properties = GetWritableProperties(record);
            var setClause = string.Join(", ", properties.Select(p => $"{p.Name} = @{p.Name}"));
            var primaryKey = GetPrimaryKeyProperty();
            var primaryKeyValue = primaryKey.GetValue(record);

            var sql = $"UPDATE {_currentTable} SET {setClause} WHERE {primaryKey.Name} = @{primaryKey.Name}";
            var parameters = CreateParameters(properties, record);

            // Добавляем первичный ключ для условия WHERE
            var primaryKeyParam = new SQLiteParameter($"@{primaryKey.Name}", primaryKeyValue);
            var paramList = parameters.ToList();
            paramList.Add(primaryKeyParam);

            _database.Query(sql, paramList.ToArray());
        }

        private void DeleteRecord(object record)
        {
            var primaryKey = GetPrimaryKeyProperty();
            var primaryKeyValue = primaryKey.GetValue(record);

            var sql = $"DELETE FROM {_currentTable} WHERE {primaryKey.Name} = @{primaryKey.Name}";
            var parameters = new SQLiteParameter[]
            {
                new SQLiteParameter($"@{primaryKey.Name}", primaryKeyValue)
            };

            _database.Query(sql, parameters);
        }

        private PropertyInfo[] GetWritableProperties(object record)
        {
            // Исключаем навигационные свойства и коллекции
            return record.GetType().GetProperties()
                .Where(p => p.CanWrite &&
                           !p.PropertyType.IsGenericType &&
                           !p.PropertyType.Namespace.StartsWith("System.Collections"))
                .ToArray();
        }

        private PropertyInfo GetPrimaryKeyProperty()
        {
            // Ищем свойство с именем, содержащим "ID" или "_ID"
            var properties = _currentType.GetProperties();

            var primaryKey = properties.FirstOrDefault(p =>
                p.Name.EndsWith("_ID") || p.Name.EndsWith("ID") ||
                p.Name.Equals("Id", StringComparison.OrdinalIgnoreCase));

            if (primaryKey == null)
            {
                // Если не нашли стандартный ID, берем первое свойство
                primaryKey = properties.First();
            }

            return primaryKey;
        }

        private SQLiteParameter[] CreateParameters(PropertyInfo[] properties, object record)
        {
            var parameters = new List<SQLiteParameter>();

            foreach (var property in properties)
            {
                var value = property.GetValue(record);
                parameters.Add(new SQLiteParameter($"@{property.Name}", value ?? DBNull.Value));
            }

            return parameters.ToArray();
        }

        private void DataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            StatusText.Text = "Данные изменены. Не забудьте сохранить!";
        }

        private void DataGrid_AddingNewItem(object sender, AddingNewItemEventArgs e)
        {
            StatusText.Text = "Добавляется новая запись...";
        }

        private void DataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            StatusText.Text = "Редактирование записи...";
        }
    }
}
