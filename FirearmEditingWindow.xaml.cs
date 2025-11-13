using RANGER.Database;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
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
using System.Windows.Shapes;
using static RANGER.Database.DataBaseModel;

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для FirearmEditingWindow.xaml
    /// </summary>
    public partial class FirearmEditingWindow : Window
    {
        private Firearm firearm;
        private DataBaseConnection db;
        private bool isNew;

        public FirearmEditingWindow(Firearm firearm, DataBaseConnection db)
        {
            InitializeComponent();

            this.firearm = firearm;
            this.db = db;
            this.isNew = firearm == null;

            if (isNew)
            {
                this.firearm = new Firearm();
                Title = "Регистрация оружия в системе";
            }
            else
            {
                Title = "Редактирование информации об оружии";
                txtName.IsEnabled = false;
                txtSerialNumber.IsEnabled = false;
                cmbCategory.IsEnabled = false;

                LoadFirearmData();
            }
        }

        private void LoadFirearmData()
        {
            txtSerialNumber.Text = firearm.FirearmSerialNumber;
            txtName.Text = firearm.Name;
            cmbCategory.Text = firearm.Category;
            cmbCondition.Text = firearm.Condition;
            dpMaitenanceDate.SelectedDate = firearm.LastMaitenanceDate;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text) ||
                string.IsNullOrWhiteSpace(txtSerialNumber.Text) ||
                string.IsNullOrWhiteSpace(cmbCategory.Text) ||
                string.IsNullOrWhiteSpace(cmbCondition.Text) ||
                dpMaitenanceDate.SelectedDate == null)
            {
                MessageBox.Show("Заполните все обязательные поля", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                firearm.FirearmSerialNumber = txtSerialNumber.Text;
                firearm.Name = txtName.Text;
                firearm.Category = cmbCategory.Text;
                firearm.Condition = cmbCondition.Text;
                firearm.LastMaitenanceDate = dpMaitenanceDate.SelectedDate.Value;

                if (isNew)
                {
                    // Вставка новой записи
                    db.Query(
                        "INSERT INTO Firearm (FirearmSerialNumber, Name, Category, Condition, LastMaitenanceDate) " +
                        "VALUES (@serialNumber, @name, @category, @condition, @lastMaitenance)",
                        new SQLiteParameter[] {
                            new SQLiteParameter("@serialNumber", firearm.FirearmSerialNumber),
                            new SQLiteParameter("@name", firearm.Name),
                            new SQLiteParameter("@category", firearm.Category),
                            new SQLiteParameter("@condition", firearm.Condition),
                            new SQLiteParameter("@lastMaitenance", firearm.LastMaitenanceDate),
                        });
                }
                else
                {
                    // Обновление существующей записи
                    // Есть такая темка, что если оружие уже внесено в базу, можно сменить только состояние или удалить
                    // Мб так реализовать...
                    db.Query(
                        "UPDATE Firearm SET Condition = @condition, LastMaitenanceDate = @lastMaitenance " +
                        "WHERE FirearmSerialNumber = @serialNumber",
                        new SQLiteParameter[] {
                            new SQLiteParameter("@condition", firearm.Condition),
                            new SQLiteParameter("@lastMaitenance", firearm.LastMaitenanceDate)
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

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
