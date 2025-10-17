using System;
using System.Collections.Generic;
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

namespace RANGER
{
    /// <summary>
    /// Логика взаимодействия для LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        private void LoginPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            UpdateVisibility();
        }

        private void UpdateVisibility()
        {
            // Ищем TextBlock с watermark в визуальном дереве
            var passwordBox = LoginPasswordBox;
            var grid = VisualTreeHelper.GetChild(passwordBox, 0) as Grid;

            if (grid != null)
            {
                var watermarkText = grid.FindName("watermarkText") as TextBlock;

                if (watermarkText != null)
                {
                    // Показывает плашку "введите ваш пароль..." только если пароль пустой
                    if (string.IsNullOrEmpty(passwordBox.Password))
                    {
                        watermarkText.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        watermarkText.Visibility = Visibility.Collapsed;
                    }
                }
            }
        }
    }
}
