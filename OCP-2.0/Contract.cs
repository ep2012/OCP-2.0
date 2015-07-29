using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace OCP_2._0
{
    public class Contract
    {
        private List<object> ids;
        private List<object> values;
        private int length;
        private int index;

        private String ELISCHEMA = @"[HEALTHCARE\eliprice]";
        private String DBOSCHEMA = "[dbo]";
        public Contract()
        {
            setUp();
        }
        private void setUp()
        {
            ids = new List<object>();
            index = 0;
            new idMaker("SELECT CONTRACT_ID FROM [REVINT]." + getSchema() + ".[OCP_Contracts] ORDER BY CLIENT", ids);
            length = ids.Count;
            if (length > 0)
            {
                setValues();
            }
            else
            {
                values = new List<object>();
            }
        }
        public void setValues()
        {
            values = new List<object>();
            new idMaker("SELECT * FROM [REVINT]." + getSchema() + ".[OCP_Contracts] WHERE CONTRACT_ID = " + getId(index), values);
        }
        public void first()
        {
            if (length == 0)
                return;
            index = 0;
            setValues();
        }
        public void increment()
        {
            if (length == 0)
                return;
            if (index == length - 1)
                index = 0;
            else
                index++;
            setValues();
        }
        public void decrement()
        {
            if (length == 0)
                return;
            if (index == 0)
                index = length - 1;
            else
                index--;
            setValues();
        }
        public String getValue(int i)
        {
            try
            {
                return values[i].ToString();
            }
            catch (Exception)
            {
                return null;
            }
        }
        private String getId(int i)
        {
            try
            {
                return ids[i].ToString();
            }
            catch(Exception)
            {
                return null;
            }
        }
        public void updateSelection(object client, object status, object peopleSoftContract, object typeOfAgreement, object department, object specialty,
            object durationOfContract, object summaryOfServices, object departmentContact, object clientLocation, object hardTerminationDate, object nextAutoRenewalDate,
            object effectiveDate, object committeeDate, object frequencyOfServices, object credentialProcessing, object mostRecentWksht, object UIBilling, object BAA,
            object consultOrProcedure, object reimbursement1, object reimbursement2, object reimbursement3, object travelReimbursement, object reimbursement1Period,
            object reimbursement2Period, object reimbursement3Period, object travelReimbursementPeriod, object individualResponsible, object fmr, object lastFMRReviewDate,
            object mfk, object timesheetInformation, object estimatedAnnualValue, object reimbursementNotes, object providers)
        {
            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";
            using (OdbcConnection dbConnection = new OdbcConnection(cxnString))
            {
                //open OdbcConnection object
                dbConnection.Open();

                OdbcCommand cmd = new OdbcCommand();

                cmd.CommandText = "{CALL [REVINT]." + getSchema() + ".[OCP_updateContract](" +
                    " ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " +
                    " ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?"+
                    " )}";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Connection = dbConnection;

                cmd.Parameters.Add("@contractId", OdbcType.Int).Value = ids[index];
                cmd.Parameters.Add("@client", OdbcType.NVarChar, 400).Value = client;
                cmd.Parameters.Add("@status", OdbcType.NVarChar, 400).Value = status;
                cmd.Parameters.Add("@peopleSoftContract", OdbcType.NVarChar, 400).Value = peopleSoftContract;
                cmd.Parameters.Add("@typeOfAgreement", OdbcType.NVarChar, 400).Value = typeOfAgreement;
                cmd.Parameters.Add("@department", OdbcType.NVarChar, 400).Value = department;
                cmd.Parameters.Add("@specialty", OdbcType.NVarChar, 400).Value = specialty;
                cmd.Parameters.Add("@durationOfContract", OdbcType.NVarChar, 400).Value = durationOfContract;
                cmd.Parameters.Add("@summaryOfServices", OdbcType.NVarChar, 4000).Value = summaryOfServices;
                cmd.Parameters.Add("@departmentContact", OdbcType.NVarChar, 400).Value = departmentContact;
                cmd.Parameters.Add("@clientLocation", OdbcType.NVarChar, 400).Value = clientLocation;
                cmd.Parameters.Add("@hardTerminationDate", OdbcType.DateTime).Value = asDateTime(hardTerminationDate);
                cmd.Parameters.Add("@nextAutoRenewalDate", OdbcType.DateTime).Value = asDateTime(nextAutoRenewalDate);
                cmd.Parameters.Add("@effectiveDate", OdbcType.DateTime).Value = asDateTime(effectiveDate);
                cmd.Parameters.Add("@committeeDate", OdbcType.DateTime).Value = asDateTime(committeeDate);
                cmd.Parameters.Add("@frequencyOfServices", OdbcType.NVarChar, 400).Value = frequencyOfServices;
                cmd.Parameters.Add("@credentialProcessing", OdbcType.NVarChar, 400).Value = credentialProcessing;
                cmd.Parameters.Add("@mostRecentWksht", OdbcType.DateTime).Value = asDateTime(mostRecentWksht);
                cmd.Parameters.Add("@UIBilling", OdbcType.Bit).Value = UIBilling;
                cmd.Parameters.Add("@BAA", OdbcType.Bit).Value = BAA;
                cmd.Parameters.Add("@consultOrProcedure", OdbcType.Int).Value = consultOrProcedure;
                cmd.Parameters.Add("@reimbursement1", OdbcType.Double).Value = asDouble(reimbursement1);
                cmd.Parameters.Add("@reimbursement2", OdbcType.Double).Value = asDouble(reimbursement2);
                cmd.Parameters.Add("@reimbursement3", OdbcType.Double).Value = asDouble(reimbursement3);
                cmd.Parameters.Add("@travelReimbursement", OdbcType.Double).Value = asDouble(travelReimbursement);
                cmd.Parameters.Add("@reimbursement1Period", OdbcType.NVarChar, 400).Value = reimbursement1Period;
                cmd.Parameters.Add("@reimbursement2Period", OdbcType.NVarChar, 400).Value = reimbursement2Period;
                cmd.Parameters.Add("@reimbursement3Period", OdbcType.NVarChar, 400).Value = reimbursement3Period;
                cmd.Parameters.Add("@travelReimbursementPeriod", OdbcType.NVarChar, 400).Value = travelReimbursementPeriod;
                cmd.Parameters.Add("@individualResponsible", OdbcType.NVarChar, 400).Value = individualResponsible;
                cmd.Parameters.Add("@fmr", OdbcType.NVarChar, 400).Value = fmr;
                cmd.Parameters.Add("@lastFMRReviewDate", OdbcType.DateTime).Value = asDateTime(lastFMRReviewDate);
                cmd.Parameters.Add("@mfk", OdbcType.NVarChar, 400).Value = mfk;
                cmd.Parameters.Add("@timesheetInformation", OdbcType.NVarChar, 400).Value = timesheetInformation;
                cmd.Parameters.Add("@estimatedAnnualValue", OdbcType.NVarChar, 400).Value = estimatedAnnualValue;
                cmd.Parameters.Add("@reimbursementNotes", OdbcType.NVarChar, 4000).Value = reimbursementNotes;
                cmd.Parameters.Add("@providers", OdbcType.NVarChar, 4000).Value = providers;
                cmd.Parameters.Add("@numRecords", OdbcType.Int);
                cmd.Parameters["@numRecords"].Direction = System.Data.ParameterDirection.ReturnValue;

                cmd.ExecuteNonQuery();

                dbConnection.Close();
            }
        }
        private object asDateTime(object dateTime)
        {
            if (dateTime == null)
                return null;
            else
                return (dateTime.ToString()).AsDBDateTime();
        }
        private object asDouble(object dblvalue)
        {
            double dblresult;
            if (dblvalue != null)
            {
                if (Double.TryParse(dblvalue.ToString(), out dblresult))
                {
                    return dblresult;
                }
                else
                {
                    return DBNull.Value;
                }

            }
            else
            {
                return DBNull.Value;
            }
        }
        public void exportSubset(String client, String department)
        {
            String sqlString = "SELECT * FROM [REVINT]." + getSchema() + ".[OCP_Contracts] WHERE 1 = 1";

            if (client != "")
            {
                sqlString += " AND CLIENT = '" + client + "'";
            }
            if (department != "")
            {
                sqlString += " AND UI_DEPARTMENT = '" + department + "'";
            }

            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";

            DataGrid dta = new DataGrid();

            //create an OdbcConnection object and connect it to the data source.
            using (OdbcConnection dbConnection = new OdbcConnection(cxnString))
            {
                //open OdbcConnection object
                dbConnection.Open();

                //Create adapter from connection and sql to obtain desired data
                OdbcDataAdapter dadapter = new OdbcDataAdapter(sqlString, dbConnection);

                //Create a table and fill it with the data from the adapter
                DataTable dtable = new DataTable();
                dadapter.Fill(dtable);

                //set the contents of the gui grid table to the data table created
                dta.ItemsSource = dtable.DefaultView;

                //Close connection
                dbConnection.Close();
            }
            export(dta);
        }
        public void exportAll(DataGrid dta)
        {
            String sqlString = "SELECT * FROM [REVINT]." + getSchema() + ".[OCP_Contracts]";

            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";

            //create an OdbcConnection object and connect it to the data source.
            using (OdbcConnection dbConnection = new OdbcConnection(cxnString))
            {
                //open OdbcConnection object
                dbConnection.Open();

                //Create adapter from connection and sql to obtain desired data
                OdbcDataAdapter dadapter = new OdbcDataAdapter(sqlString, dbConnection);

                //Create a table and fill it with the data from the adapter
                DataTable dtable = new DataTable();
                dadapter.Fill(dtable);

                //set the contents of the gui grid table to the data table created
                dta.ItemsSource = dtable.DefaultView;

                //Close connection
                dbConnection.Close();
            }
            export(dta);
        }

        public void export(DataGrid dta)
        {
            Microsoft.Office.Interop.Excel.Application app = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                workbook = app.Workbooks.Add();
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;

                // Put Column Header into excel work sheet
                for (int i = 0; i < dta.Columns.Count; i++)
                {
                    worksheet.Range["A1"].Offset[0, i].Value = dta.Columns[i].Header;
                }

                Microsoft.Office.Interop.Excel.Range firstRow = (Microsoft.Office.Interop.Excel.Range)worksheet.Rows[1];
                firstRow.Cells.Interior.ColorIndex = 36;

                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                firstRow.EntireRow.Font.Bold = true;

                BorderAround(firstRow, 0);

                for (int rowIndex = 0; rowIndex < dta.Items.Count; rowIndex++)
                    for (int columnIndex = 0; columnIndex < dta.Columns.Count; columnIndex++)
                    {
                        worksheet.Range["A2"].Offset[rowIndex, columnIndex].Value =
                          (dta.Items[rowIndex] as DataRowView).Row.ItemArray[columnIndex].ToString();
                    }
                worksheet.Columns.AutoFit();
                app.Visible = true;
            }
            catch (Exception)
            {
                Console.Write("Error");
            }
        }
        private void BorderAround(Microsoft.Office.Interop.Excel.Range range, int color)
        {
            Microsoft.Office.Interop.Excel.Borders borders = range.Borders;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders.Color = color;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders = null;
        }
        public void addNew(object client, object status, object peopleSoftContract, object typeOfAgreement, object department, object specialty,
            object durationOfContract, object summaryOfServices, object departmentContact, object clientLocation, object hardTerminationDate, object nextAutoRenewalDate,
            object effectiveDate, object committeeDate, object frequencyOfServices, object credentialProcessing, object mostRecentWksht, object UIBilling, object BAA,
            object consultOrProcedure, object reimbursement1, object reimbursement2, object reimbursement3, object travelReimbursement, object reimbursement1Period,
            object reimbursement2Period, object reimbursement3Period, object travelReimbursementPeriod, object individualResponsible, object fmr, object lastFMRReviewDate,
            object mfk, object timesheetInformation, object estimatedAnnualValue, object reimbursementNotes, object providers)
        {
            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";
            using (OdbcConnection dbConnection = new OdbcConnection(cxnString))
            {
                //open OdbcConnection object
                dbConnection.Open();

                OdbcCommand cmd = new OdbcCommand();

                cmd.CommandText = "{CALL [REVINT]." + getSchema() + ".[OCP_insertContract](" +
                    " ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " +
                    " ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?" +
                    " )}";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Connection = dbConnection;

                cmd.Parameters.Add("@client", OdbcType.NVarChar, 400).Value = client;
                cmd.Parameters.Add("@status", OdbcType.NVarChar, 400).Value = status;
                cmd.Parameters.Add("@peopleSoftContract", OdbcType.NVarChar, 400).Value = peopleSoftContract;
                cmd.Parameters.Add("@typeOfAgreement", OdbcType.NVarChar, 400).Value = typeOfAgreement;
                cmd.Parameters.Add("@department", OdbcType.NVarChar, 400).Value = department;
                cmd.Parameters.Add("@specialty", OdbcType.NVarChar, 400).Value = specialty;
                cmd.Parameters.Add("@durationOfContract", OdbcType.NVarChar, 400).Value = durationOfContract;
                cmd.Parameters.Add("@summaryOfServices", OdbcType.NVarChar, 4000).Value = summaryOfServices;
                cmd.Parameters.Add("@departmentContact", OdbcType.NVarChar, 400).Value = departmentContact;
                cmd.Parameters.Add("@clientLocation", OdbcType.NVarChar, 400).Value = clientLocation;
                cmd.Parameters.Add("@hardTerminationDate", OdbcType.DateTime).Value = asDateTime(hardTerminationDate);
                cmd.Parameters.Add("@nextAutoRenewalDate", OdbcType.DateTime).Value = asDateTime(nextAutoRenewalDate);
                cmd.Parameters.Add("@effectiveDate", OdbcType.DateTime).Value = asDateTime(effectiveDate);
                cmd.Parameters.Add("@committeeDate", OdbcType.DateTime).Value = asDateTime(committeeDate);
                cmd.Parameters.Add("@frequencyOfServices", OdbcType.NVarChar, 400).Value = frequencyOfServices;
                cmd.Parameters.Add("@credentialProcessing", OdbcType.NVarChar, 400).Value = credentialProcessing;
                cmd.Parameters.Add("@mostRecentWksht", OdbcType.DateTime).Value = asDateTime(mostRecentWksht);
                cmd.Parameters.Add("@UIBilling", OdbcType.Bit).Value = UIBilling;
                cmd.Parameters.Add("@BAA", OdbcType.Bit).Value = BAA;
                cmd.Parameters.Add("@consultOrProcedure", OdbcType.Int).Value = consultOrProcedure;
                cmd.Parameters.Add("@reimbursement1", OdbcType.Double).Value = asDouble(reimbursement1);
                cmd.Parameters.Add("@reimbursement2", OdbcType.Double).Value = asDouble(reimbursement2);
                cmd.Parameters.Add("@reimbursement3", OdbcType.Double).Value = asDouble(reimbursement3);
                cmd.Parameters.Add("@travelReimbursement", OdbcType.Double).Value = asDouble(travelReimbursement);
                cmd.Parameters.Add("@reimbursement1Period", OdbcType.NVarChar, 400).Value = reimbursement1Period;
                cmd.Parameters.Add("@reimbursement2Period", OdbcType.NVarChar, 400).Value = reimbursement2Period;
                cmd.Parameters.Add("@reimbursement3Period", OdbcType.NVarChar, 400).Value = reimbursement3Period;
                cmd.Parameters.Add("@travelReimbursementPeriod", OdbcType.NVarChar, 400).Value = travelReimbursementPeriod;
                cmd.Parameters.Add("@individualResponsible", OdbcType.NVarChar, 400).Value = individualResponsible;
                cmd.Parameters.Add("@fmr", OdbcType.NVarChar, 400).Value = fmr;
                cmd.Parameters.Add("@lastFMRReviewDate", OdbcType.DateTime).Value = asDateTime(lastFMRReviewDate);
                cmd.Parameters.Add("@mfk", OdbcType.NVarChar, 400).Value = mfk;
                cmd.Parameters.Add("@timesheetInformation", OdbcType.NVarChar, 400).Value = timesheetInformation;
                cmd.Parameters.Add("@estimatedAnnualValue", OdbcType.NVarChar, 400).Value = estimatedAnnualValue;
                cmd.Parameters.Add("@reimbursementNotes", OdbcType.NVarChar, 4000).Value = reimbursementNotes;
                cmd.Parameters.Add("@providers", OdbcType.NVarChar, 4000).Value = providers;
                cmd.Parameters.Add("@numRecords", OdbcType.Int);
                cmd.Parameters["@numRecords"].Direction = System.Data.ParameterDirection.ReturnValue;

                cmd.ExecuteNonQuery();

                dbConnection.Close();
            }
        }
        public int getLength()
        {
            return length;
        }

        public int getIndex()
        {
            return index;
        }
        public void deleteSelection()
        {
            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";
            using (OdbcConnection dbConnection = new OdbcConnection(cxnString))
            {
                //open OdbcConnection object
                dbConnection.Open();

                OdbcCommand cmd = new OdbcCommand();

                cmd.CommandText = "{CALL [REVINT]." + getSchema() + ".[OCP_deleteContract]( ?, ? )}";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Connection = dbConnection;

                cmd.Parameters.Add("@contractId", OdbcType.Int).Value = ids[index];
                cmd.Parameters.Add("@numRecords", OdbcType.Int);
                cmd.Parameters["@numRecords"].Direction = System.Data.ParameterDirection.ReturnValue;

                cmd.ExecuteNonQuery();

                dbConnection.Close();
            }
            setUp();
        }
        public String getSchema()
        {
            return ELISCHEMA;
        }

        //Need select stored procedure
        public String filter(String client, String department, String [] selectedStatuses, bool isLessThanNinetyChecked)
        {
            ids = new List<object>();
            index = 0;
            String sqlString = " FROM [REVINT]." + getSchema() + ".[OCP_Contracts] WHERE 1 = 1";

            if (client != "")
            {
                sqlString += " AND CLIENT = '" + client.Replace("'","''".ToString()) + "'";
            }
            if (department != "")
            {
                sqlString += " AND UI_DEPARTMENT = '" + department + "'";
            }
            
            if (selectedStatuses.Length > 0)
            {
                sqlString += " AND ( STATUS = '" +selectedStatuses[0]+"'";

                for (int i = 1; i < selectedStatuses.Length; i++)
                {
                    sqlString += " OR STATUS = '" + selectedStatuses[i] + "'";
                }
                sqlString += ")";
            }

            if(isLessThanNinetyChecked)
            {
                sqlString += " AND (NEXT_AUTO_RENEWAL_DATE > '" + DateTime.Today.ToShortDateString() + "' AND NEXT_AUTO_RENEWAL_DATE < '" + DateTime.Today.AddDays(90).ToShortDateString() + "')";
            }

            new idMaker("SELECT CONTRACT_ID" + sqlString + " ORDER BY CLIENT", ids);

            length = ids.Count;
            if (length > 0)
            {
                setValues();
            }
            else
            {
                values = new List<object>();
            }
            return sqlString;
        }
    }

    public class idMaker
    {
        public idMaker(String sqlCmdID, List<object> ids)
        {
            object[] objID = new object[40];

            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";

            using (OdbcConnection connectionID = new OdbcConnection(cxnString))
            {
                OdbcCommand commandID = new OdbcCommand(@sqlCmdID, connectionID);

                connectionID.Open();

                OdbcDataReader reader = commandID.ExecuteReader();

                while (reader.Read())
                {
                    int numCols = reader.GetValues(objID);

                    for (int i = 0; i < numCols; i++)
                    {
                        ids.Add(objID[i]);
                    }
                } 
            }
        }
    }
    public class comboMaker
    {
        public comboMaker(String sqlCmd, String sqlCmdID, Dictionary<String, String> dictionary, ObservableCollection<string> comboClass)
        {
            object[] obj = new object[10];
            object[] objID = new object[10];

            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";

            OdbcConnection connection = new OdbcConnection(cxnString);

            using (OdbcConnection connectionID = new OdbcConnection(cxnString))
            {
                OdbcCommand command = new OdbcCommand(sqlCmd, connection);
                OdbcCommand commandID = new OdbcCommand(sqlCmdID, connectionID);

                connection.Open();
                connectionID.Open();

                OdbcDataReader reader = command.ExecuteReader();
                OdbcDataReader reader2 = commandID.ExecuteReader();

                while (reader.Read() && reader2.Read())
                {
                    int numCols = reader.GetValues(obj);
                    numCols = reader2.GetValues(objID);

                    for (int i = 0; i < numCols; i++)
                    {
                        dictionary.Add(obj[i].ToString(), objID[i].ToString());
                        comboClass.Add(obj[i].ToString());
                    }
                }
            }
        }
    }
}
public static class DBValueExtensions
{
    public static object AsDBDateTime(this string s)
    {
        var str = s;
        if (null != str) { str = str.Trim(); }

        if (string.IsNullOrEmpty(str))
        {
            return DBNull.Value;
        }
        else
        {
            DateTime time;
            if (DateTime.TryParse(str, out time))
            {
                return time;
            }
            else
            {
                return DBNull.Value;
            }
        }
    }
}