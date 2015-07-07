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

namespace OCP_2._0
{
    /// <summary>
    /// Interaction logic for Menu.xaml
    /// </summary>
    public partial class Menu : Window
    {
        public Menu()
        {
            InitializeComponent();
            lblDate.Content = DateTime.Today.ToLongDateString();
            lblUsername.Content = "Logged in as: " + Environment.UserName;
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnAddNewUser_Click(object sender, RoutedEventArgs e)
        {
            AddUser win = new AddUser();
            win.Left = this.Left;
            win.Top = this.Top;
            win.Show();
            this.Close();
        }

        private void btnOneContractPerPage_Click(object sender, RoutedEventArgs e)
        {
            Contracts win = new Contracts();
            win.Left = this.Left;
            win.Top = this.Top;
            win.Show();
            this.Close();
        }
    }
}
