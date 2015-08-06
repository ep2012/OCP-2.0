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
using System.Windows.Shapes;

namespace OCP_2._0
{
    /// <summary>
    /// Interaction logic for AddUser.xaml
    /// </summary>
    public partial class AddUser : Window
    {
        int SHORT_VARCHAR_LENGTH = 400;

        public AddUser()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            Menu win = new Menu();
            win.Left = this.Left;
            win.Top = this.Top;
            win.Show();
            this.Close();
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (isFormValid())
            {
                if (addUser() == "1")
                {
                    Menu win = new Menu();
                    win.Left = this.Left;
                    win.Top = this.Top;
                    win.Show();
                    this.Close();
                }
                else
                {
                    //not successful
                    var dialogBox = MessageBox.Show("A user with the given Healthcare Id already exists. " +
                        "Create a different user or contact a programmer for more assistance.",
                        "Unsuccessful Database Insertion");
                }
            }
            else
            {
                var dialogBox = MessageBox.Show("The given information is not valid to insert into the database. " +
                        "Check that the name and Healthcare Id fields are filled properly. " + 
                        "If this problem persists, contact a programmer for further assistance.",
                        "Invalid Entry");
            }
        }
        private bool isFormValid()
        {
            if ((txtHealthcareID.Text == "" && txtHealthcareID.Text.Length < SHORT_VARCHAR_LENGTH) || (txtName.Text == "" && txtName.Text.Length < SHORT_VARCHAR_LENGTH))
            {
                return false;
            }
            return true;
        }
        private String addUser()
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

                    cmd.CommandText = "{CALL [REVINT]." + contract.getSchema() + ".[OCP_addUser]( ?, ?, ?, ? )}";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Connection = dbConnection;

                    cmd.Parameters.Add("@hawkId", OdbcType.NVarChar, 400).Value = txtHealthcareID.Text;
                    cmd.Parameters.Add("@administrator", OdbcType.Bit).Value = chkAdmin.IsChecked;
                    cmd.Parameters.Add("@name", OdbcType.NVarChar, 400).Value = txtName.Text;
                    cmd.Parameters.Add("@numRecords", OdbcType.Int);
                    cmd.Parameters["@numRecords"].Direction = System.Data.ParameterDirection.ReturnValue;

                    cmd.ExecuteNonQuery();

                    dbConnection.Close();

                    return cmd.Parameters["@numRecords"].Value.ToString();
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
