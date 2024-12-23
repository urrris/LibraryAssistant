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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Npgsql;


namespace LibraryAssistant
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void DatabaseConnectionConnectButton_Click(object sender, RoutedEventArgs e)
        {
            TextBox host = (TextBox)this.FindName("DatabaseConnectionHostInput");
            TextBox port = (TextBox)this.FindName("DatabaseConnectionPortInput");
            TextBox user = (TextBox)this.FindName("DatabaseConnectionUserInput");
            TextBox password = (TextBox)this.FindName("DatabaseConnectionPasswordInput");
            TextBox databaseName = (TextBox)this.FindName("DatabaseConnectionDatabaseNameInput");

            if (host.Text == "" || port.Text == "" || user.Text == "" || password.Text == "" || databaseName.Text == "") {
                TextBox errorMessage = (TextBox)this.FindName("DatabaseConnectionErrorMessage");
                errorMessage.Foreground = Brushes.Red;
                errorMessage.Text = "Все поля должны быть заполнены";
                return;
            }

            string connectionString = String.Format("Host={0};Port={1};Database={2};Username={3};Password={4};", host.Text, port.Text, databaseName.Text, user.Text, password.Text);
            NpgsqlConnection connection = new NpgsqlConnection(connectionString);

            try {
                connection.Open();
                connection.Close();
                ApplicationWindow window = new ApplicationWindow(connection);
                window.Show();
                this.Close();
            }
            catch {
                connection.Close();
                TextBox errorMessage = (TextBox)this.FindName("DatabaseConnectionErrorMessage");
                errorMessage.Foreground = Brushes.Red;
                errorMessage.Text = "Не удалось подключиться к БД";
            }
        }
    }
}
