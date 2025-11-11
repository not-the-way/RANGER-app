using RANGER.Database;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using static RANGER.Database.DataBaseModel;

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для EmployeeEditingWindow.xaml
    /// </summary>
    public partial class EmployeeEditingWindow : Window
    {
        private Employee employee;
        private DataBaseConnection db;
        private bool isNew;

        public EmployeeEditingWindow(Employee employee, DataBaseConnection db)
        {
            InitializeComponent();

            this.employee = employee;
            this.db = db;
            this.isNew = employee == null;

            if (isNew)
            {
                this.employee = new Employee();
                Title = "Добавление сотрудника";
            }
            else
            {
                Title = "Редактирование сотрудника";
                LoadEmployeeData();
            }
        }

        private void LoadEmployeeData()
        {
            txtFullName.Text = employee.FullName;
            dpEmploymentDate.SelectedDate = employee.EmploymentDate;
            txtLogin.Text = employee.Login;
            txtPassword.Password = employee.Password;
            cmbAccessLevel.Text = employee.AccessLevel;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFullName.Text) ||
                string.IsNullOrWhiteSpace(txtLogin.Text) ||
                dpEmploymentDate.SelectedDate == null)
            {
                MessageBox.Show("Заполните все обязательные поля", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                employee.FullName = txtFullName.Text;
                employee.EmploymentDate = dpEmploymentDate.SelectedDate.Value;
                employee.Login = txtLogin.Text;
                employee.AccessLevel = cmbAccessLevel.Text;

                if (!string.IsNullOrEmpty(txtPassword.Password))
                {
                    employee.Password = txtPassword.Password;
                }

                if (isNew)
                {
                    // Вставка новой записи
                    db.Query(
                        "INSERT INTO Employee (FullName, EmploymentDate, Login, Password, AccessLevel) " +
                        "VALUES (@fullName, @empDate, @login, @password, @accessLevel)",
                        new SQLiteParameter[] {
                            new SQLiteParameter("@fullName", employee.FullName),
                            new SQLiteParameter("@empDate", employee.EmploymentDate),
                            new SQLiteParameter("@login", employee.Login),
                            new SQLiteParameter("@password", employee.Password),
                            new SQLiteParameter("@accessLevel", employee.AccessLevel)
                        });
                }
                else
                {
                    // Обновление существующей записи
                    if (string.IsNullOrEmpty(txtPassword.Password))
                    {
                        db.Query(
                            "UPDATE Employee SET FullName = @fullName, EmploymentDate = @empDate, " +
                            "Login = @login, AccessLevel = @accessLevel WHERE Employee_ID = @id",
                            new SQLiteParameter[] {
                                new SQLiteParameter("@fullName", employee.FullName),
                                new SQLiteParameter("@empDate", employee.EmploymentDate),
                                new SQLiteParameter("@login", employee.Login),
                                new SQLiteParameter("@accessLevel", employee.AccessLevel),
                                new SQLiteParameter("@id", employee.Employee_ID)
                            });
                    }
                    else
                    {
                        db.Query(
                            "UPDATE Employee SET FullName = @fullName, EmploymentDate = @empDate, " +
                            "Login = @login, Password = @password, AccessLevel = @accessLevel WHERE Employee_ID = @id",
                            new SQLiteParameter[] {
                                new SQLiteParameter("@fullName", employee.FullName),
                                new SQLiteParameter("@empDate", employee.EmploymentDate),
                                new SQLiteParameter("@login", employee.Login),
                                new SQLiteParameter("@password", employee.Password),
                                new SQLiteParameter("@accessLevel", employee.AccessLevel),
                                new SQLiteParameter("@id", employee.Employee_ID)
                            });
                    }
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
