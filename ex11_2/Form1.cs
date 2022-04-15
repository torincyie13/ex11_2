using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ADOX;
using System.Data.OleDb;
using System.IO;

namespace ex11_2
{
    public partial class frmImport : Form
    {
        OleDbConnection weatherConnection; 
        OleDbCommand weatherCommand; 
        OleDbDataAdapter weatherAdapter; 
        DataTable weatherTable;
        public frmImport()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            Catalog weatherDatabase = new Catalog(); 
            // Table databaseTable;
            string databaseFile = Application.StartupPath + "\\Sea02DB.accdb";
            string dataFile = Application.StartupPath + "\\Sea02.txt"; 
            string dataYear = "2002"; 
            // create empty database
            //try
            //{
            //    weatherDatabase.Create("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + databaseFile); 
            //    // create table
            //    databaseTable = new Table();
            //    databaseTable.Name = "Weather";
            //    // add fields
            //    databaseTable.Columns.Append("WeatherDate", DataTypeEnum.adDate, 20);
            //    databaseTable.Columns.Append("HighTemp", DataTypeEnum.adSingle, 10); 
            //    databaseTable.Columns.Append("LowTemp", DataTypeEnum.adSingle, 10);
            //    databaseTable.Columns.Append("Precip", DataTypeEnum.adSingle, 10); 
            //    databaseTable.Columns.Append("Comment", DataTypeEnum.adWChar, 50);
            //    databaseTable.Columns["HighTemp"].Attributes = ColumnAttributesEnum.adColNullable;
            //    databaseTable.Columns["LowTemp"].Attributes = ColumnAttributesEnum.adColNullable;
            //    databaseTable.Columns["Precip"].Attributes = ColumnAttributesEnum.adColNullable;
            //    databaseTable.Columns["Comment"].Attributes = ColumnAttributesEnum.adColNullable; 
            //    // primary key
            //    databaseTable.Keys.Append("PK_Weather", KeyTypeEnum.adKeyPrimary, "WeatherDate", null, null);
            //    databaseTable.Indexes.Append("WeatherDate", "WeatherDate");
            //    weatherDatabase.Tables.Append(databaseTable);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Error Creating Database",
            //    MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //finally
            //{
            //    weatherDatabase = null;
            //}
            // connect to database
            weatherConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\thawkins022713\\source\\repos\\ex11_2\\ex11_2\\bin\\Debug\\Sea02.accdb"); 
            weatherConnection.Open();
            weatherCommand = new OleDbCommand("SELECT * FROM Weather ORDER BY WeatherDate", weatherConnection);
            weatherAdapter = new OleDbDataAdapter();
            weatherAdapter.SelectCommand = weatherCommand; 
            weatherTable = new DataTable(); weatherAdapter.Fill(weatherTable); 
            // fill dates and other fields
            StreamReader inputFile = new StreamReader(dataFile); 
            // skip first line
            inputFile.ReadLine();
            DateTime tableDate = new DateTime(Convert.ToInt32(dataYear), 1, 1);
            DateTime nextYear = new DateTime(Convert.ToInt32(dataYear) + 1, 1, 1); 
            int nDays = (nextYear - tableDate).Days; 
            DataRow newRow; 
            string myLine, myVariable;
            int cl; 
            for (int n = 1; !inputFile.EndOfStream && n <= nDays; n++)
            {
                newRow = weatherTable.NewRow(); newRow["WeatherDate"] = tableDate; 
                // read data line and parse out four fields
                myLine = inputFile.ReadLine(); 
                // high temperature
                cl = myLine.IndexOf(","); 
                myVariable = myLine.Substring(0, cl);
                newRow["HighTemp"] = Convert.ToSingle(myVariable); 
                myLine = myLine.Substring(cl + 1, myLine.Length - cl - 1).Trim(); 
                // low temperature
                cl = myLine.IndexOf(","); 
                myVariable = myLine.Substring(0, cl); 
                newRow["LowTemp"] = Convert.ToSingle(myVariable); 
                myLine = myLine.Substring(cl + 1, myLine.Length - cl - 1).Trim(); 
                // precip
                cl = myLine.IndexOf(","); ; 
                myVariable = myLine.Substring(0, cl);
                if (!myVariable.Equals("")) 
                    newRow["Precip"] = Convert.ToSingle(myVariable);
                else newRow["Precip"] = 0.0;
                myLine = myLine.Substring(cl + 1, myLine.Length - cl - 1).Trim();
                // comment
                newRow["Comment"] = myLine;
                weatherTable.Rows.Add(newRow); 
                tableDate += new TimeSpan(1, 0, 0, 0);
            }
            inputFile.Close(); 
            // save database and dispose of objects
            OleDbCommandBuilder weatherAdapterCommand = new OleDbCommandBuilder(weatherAdapter);
            weatherAdapter.Update(weatherTable); 
            weatherConnection.Close();
            weatherConnection.Dispose(); 
            weatherCommand.Dispose();
            weatherAdapter.Dispose(); 
            weatherTable.Dispose();
            weatherAdapterCommand.Dispose();
        }

    }
}
