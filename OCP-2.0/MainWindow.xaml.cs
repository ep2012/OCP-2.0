using System;
using System.Collections.Generic;
using System.Data.Odbc;
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

namespace OCP_2._0
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            if (!isRegisteredUser())
            {
                this.Close();
            }
        }
        private void btnAccept_Click(object sender, RoutedEventArgs e)
        {
            Menu win = new Menu();
            win.Left = this.Left;
            win.Top = this.Top;
            win.Show();

            this.Close();
        }

        private void btnDecline_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private bool isRegisteredUser()
        {
            try
            {
                Contract contract = new Contract();
                String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";
                using (OdbcConnection dbConnection = new OdbcConnection(cxnString))
                {
                    //open OdbcConnection object
                    dbConnection.Open();

                    OdbcCommand cmd = new OdbcCommand();

                    cmd.CommandText = "{CALL [REVINT]." + contract.getSchema() + ".[OCP_selectNumberOfUsers]( ?, ? )}";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Connection = dbConnection;

                    cmd.Parameters.Add("@hawkId", OdbcType.NVarChar, 400).Value = Environment.UserName;
                    cmd.Parameters.Add("@numUsers", OdbcType.Int);
                    cmd.Parameters["@numUsers"].Direction = System.Data.ParameterDirection.ReturnValue;

                    cmd.ExecuteNonQuery();

                    dbConnection.Close();

                    return cmd.Parameters["@numUsers"].Value.ToString().Equals("1");
                }
            }
            catch(Exception)
            {
                return false;
            }
        }


    }
}
