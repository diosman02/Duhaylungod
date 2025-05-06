using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Duhaylungod
{
    public partial class Logs: Form
    {
        public Logs()
        {
            InitializeComponent();
            LoadLogsData();
        }

        private void LoadLogsData()
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOKDB.xlsx");
            Worksheet sheet = workbook.Worksheets[1];

            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("Name");
            dataTable.Columns.Add("Saying");
            dataTable.Columns.Add("Date");
            dataTable.Columns.Add("Time");

            for (int rowIndex = 2; rowIndex <= sheet.LastRow; rowIndex++)
            {
                DataRow row = dataTable.NewRow();
                row["Name"] = sheet[rowIndex, 1].Value;
                row["Saying"] = sheet[rowIndex, 2].Value;
                row["Date"] = sheet[rowIndex, 3].Value;
                row["Time"] = sheet[rowIndex, 4].Value;
                dataTable.Rows.Add(row);
            }
            dataGridView1.DataSource = dataTable;
        }

    }
}
