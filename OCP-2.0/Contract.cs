using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OCP_2._0
{
    public class Contract
    {
        private List<object> ids;
        private List<object> values;
        private int length;
        private int index;
        public Contract()
        {
            ids = new List<object>();
            index = 0;
            new idMaker("SELECT CONTRACT_ID FROM [REVINT].[dbo].[OCP_Contracts] ORDER BY CONTRACT_ID", ids);
            length = ids.Count;
            if (length > 0)
            {
                setValues();
            }
        }
        private void setValues()
        {
            values = new List<object>();
            new idMaker("SELECT * FROM [REVINT].[dbo].[OCP_Contracts] WHERE CONTRACT_ID = " + ids[index], values);
        }
        public void increment()
        {
            if (index == length - 1)
                index = 0;
            else
                index++;
            setValues();
        }
        public void decrement()
        {
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
    }

    public class idMaker
    {
        public idMaker(String sqlCmdID, List<object> ids)
        {
            object[] objID = new object[40];

            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";

            using (OdbcConnection connectionID = new OdbcConnection(cxnString))
            {
                OdbcCommand commandID = new OdbcCommand(sqlCmdID, connectionID);

                connectionID.Open();

                OdbcDataReader reader = commandID.ExecuteReader();

                if (reader.Read())
                {
                    do
                    {
                        int numCols = reader.GetValues(objID);

                        for (int i = 0; i < numCols; i++)
                        {
                            ids.Add(objID[i]);
                        }
                    } while (reader.Read());
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

                if (reader.Read() && reader2.Read())
                {
                    do
                    {
                        int numCols = reader.GetValues(obj);
                        numCols = reader2.GetValues(objID);

                        for (int i = 0; i < numCols; i++)
                        {
                            dictionary.Add(obj[i].ToString(), objID[i].ToString());
                            comboClass.Add(obj[i].ToString());
                        }
                    } while (reader.Read() && reader2.Read());
                }
            }

        }
    }
}
