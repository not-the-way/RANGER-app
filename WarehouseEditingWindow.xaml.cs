using RANGER.Database;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using static RANGER.Database.DataBaseModel;

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для WarehouseEditingWindow.xaml
    /// </summary>
    public partial class WarehouseEditingWindow : Window
    {
        private Warehouse warehouse;
        private DataBaseConnection db;
        private bool isNew;

        private List<TypesForWarehouse> warehouseTypes;
        public WarehouseEditingWindow(Warehouse warehouse, DataBaseConnection db)
        {
            InitializeComponent();
            this.warehouse = warehouse;
            this.db = db;
            this.isNew = warehouse == null;

            if (isNew)
            {
                this.warehouse = new Warehouse();
                Title = "Добавление предмета на склад";
            }
            else
            {
                Title = "Редактирование предмета на складе";
            }

            LoadComboBoxData();
            LoadWarehouseData();
        }

        private void LoadComboBoxData()
        {
            try
            {
                // Загрузка типов склада
                warehouseTypes = db.ExecuteQuery<TypesForWarehouse>("SELECT Type FROM TypesForWarehouse");
                cmbWarehouseType.ItemsSource = warehouseTypes;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки типов склада: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadWarehouseData()
        {
            if (!isNew)
            {
                txtItemName.Text = warehouse.ItemName;
                dpDeliveryDate.SelectedDate = warehouse.DeliveryDate;
                txtQuantity.Text = warehouse.Quantity.ToString();
                
                // Установка типа склада
                if (!string.IsNullOrEmpty(warehouse.Warehouse_Type))
                {
                    var warehouseType = warehouseTypes.FirstOrDefault(wt => wt.Type == warehouse.Warehouse_Type);
                    cmbWarehouseType.SelectedItem = warehouseType;
                }
            }
            else
            {
                // Установка значений по умолчанию для новой записи
                dpDeliveryDate.SelectedDate = DateTime.Now;
                txtQuantity.Text = "1";
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInput())
                return;

            try
            {
                // Получаем выбранный тип склада
                var selectedType = cmbWarehouseType.SelectedItem as TypesForWarehouse;
                if (selectedType == null)
                {
                    MessageBox.Show("Выберите тип склада", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Обновляем данные склада
                warehouse.ItemName = txtItemName.Text.Trim();
                warehouse.DeliveryDate = dpDeliveryDate.SelectedDate.Value;
                warehouse.Warehouse_Type = selectedType.Type;
                warehouse.Quantity = int.Parse(txtQuantity.Text);

                if (isNew)
                {
                    // Вставка новой записи
                    db.Query(
                        "INSERT INTO Warehouse (ItemName, DeliveryDate, Warehouse_Type, Quantity) " +
                        "VALUES (@itemName, @deliveryDate, @warehouseType, @quantity)",
                        new SQLiteParameter[] {
                            new SQLiteParameter("@itemName", warehouse.ItemName),
                            new SQLiteParameter("@deliveryDate", warehouse.DeliveryDate),
                            new SQLiteParameter("@warehouseType", warehouse.Warehouse_Type),
                            new SQLiteParameter("@quantity", warehouse.Quantity)
                        });
                }
                else
                {
                    // Обновление существующей записи
                    db.Query(
                        "UPDATE Warehouse SET ItemName = @itemName, DeliveryDate = @deliveryDate, " +
                        "Warehouse_Type = @warehouseType, Quantity = @quantity " +
                        "WHERE Item_ID = @id",
                        new SQLiteParameter[] {
                            new SQLiteParameter("@itemName", warehouse.ItemName),
                            new SQLiteParameter("@deliveryDate", warehouse.DeliveryDate),
                            new SQLiteParameter("@warehouseType", warehouse.Warehouse_Type),
                            new SQLiteParameter("@quantity", warehouse.Quantity),
                            new SQLiteParameter("@id", warehouse.Item_ID)
                        });
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
            // Проверка названия предмета
            if (string.IsNullOrWhiteSpace(txtItemName.Text))
            {
                MessageBox.Show("Введите название предмета", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                txtItemName.Focus();
                return false;
            }

            // Проверка даты поставки
            if (dpDeliveryDate.SelectedDate == null)
            {
                MessageBox.Show("Выберите дату поставки", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                dpDeliveryDate.Focus();
                return false;
            }

            // Проверка на поставку в будущем (почему бы и нет?)
            if (dpDeliveryDate.SelectedDate > DateTime.Now)
            {
                var result = MessageBox.Show("Дата поставки в будущем. Продолжить сохранение?",
                    "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No)
                    return false;
            }

            // Проверка типа склада
            if (cmbWarehouseType.SelectedItem == null)
            {
                MessageBox.Show("Выберите тип склада", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                cmbWarehouseType.Focus();
                return false;
            }

            // Проверка количества
            if (string.IsNullOrWhiteSpace(txtQuantity.Text))
            {
                MessageBox.Show("Введите количество", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                txtQuantity.Focus();
                return false;
            }

            if (!int.TryParse(txtQuantity.Text, out int quantity) || quantity < 0)
            {
                MessageBox.Show("Количество должно быть положительным числом", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                txtQuantity.Focus();
                txtQuantity.SelectAll();
                return false;
            }

            return true;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void txtQuantity_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Разрешаем только цифры
            if (!char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
            }
        }

        // Обработчик для Paste события (запрет вставки нечисловых значений)
        private void txtQuantity_PreviewExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            if (e.Command == ApplicationCommands.Paste)
            {
                e.Handled = true;
            }
        }
    }
}
