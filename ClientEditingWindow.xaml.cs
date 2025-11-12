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

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для ClientEditingWindow.xaml
    /// </summary>
    public partial class ClientEditingWindow : Window
    {
        private bool isNew;
        private Client client;
        private DataBaseConnection db;

        public ClientEditingWindow(Client client, DataBaseConnection db)
        {
            InitializeComponent();

            this.client = client;
            this.db = db;
            this.isNew = client == null;

            if (isNew)
            {
                this.client = new Client();
                Title = "Добавление клиента";
            }
            else
            {
                Title = "Редактирование клиента";
                LoadClientData();
            }
        }
        private void LoadClientData()
        {
            txtFullName.Text = client.FullName;
            txtPhoneNumber.Text = client.PhoneNumber;
            txtPassport.Text = client.Passport;
            chkIsMature.IsChecked = client.IsMature;
            chkReleaseOfLiability.IsChecked = client.ReleaseOfLiability;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFullName.Text) ||
               string.IsNullOrWhiteSpace(txtPassport.Text) ||
               string.IsNullOrWhiteSpace(txtPhoneNumber.Text))
            {
                MessageBox.Show("Заполните все обязательные поля", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                client.FullName = txtFullName.Text;
                client.PhoneNumber = txtPhoneNumber.Text;
                client.Passport = txtPassport.Text;
                client.IsMature = (bool)chkIsMature.IsChecked;
                client.ReleaseOfLiability = (bool)chkReleaseOfLiability.IsChecked;

                if (isNew)
                {
                    // Вставка новой записи
                    db.Query(
                        "INSERT INTO Client (FullName, PhoneNumber, Passport, IsMature, ReleaseOfLiability) " +
                        "VALUES (@fullName, @phoneNum, @passport, @isMature, @release)",
                        new SQLiteParameter[] {
                            new SQLiteParameter("@fullName", client.FullName),
                            new SQLiteParameter("@phoneNum", client.PhoneNumber),
                            new SQLiteParameter("@passport", client.Passport),
                            new SQLiteParameter("@isMature", client.IsMature),
                            new SQLiteParameter("@release", client.ReleaseOfLiability)
                        });
                }
                else
                {
                    // Обновление существующей записи
                    db.Query(
                        "UPDATE Client SET FullName = @fullName, PhoneNumber = @phoneNum, " +
                        "Passport = @passport, IsMature = @isMature, ReleaseOfLiability = @release WHERE Client_ID = @id",
                        new SQLiteParameter[] {
                            new SQLiteParameter("@fullName", client.FullName),
                            new SQLiteParameter("@phoneNum", client.PhoneNumber),
                            new SQLiteParameter("@passport", client.Passport),
                            new SQLiteParameter("@isMature", client.IsMature),
                            new SQLiteParameter("@release", client.ReleaseOfLiability),
                            new SQLiteParameter("@id", client.Client_ID)
                        });
                }
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
