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
    public partial class Inactive: Form
    {
        public Inactive()
        {
            InitializeComponent();
            LoadInactiveData();
        }
        private void LoadInactiveData()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();


            DataRow[] inactiveRows = dt.Select("Status = '0'");
            DataTable inactiveTable = inactiveRows.CopyToDataTable();
            dataGridView1.DataSource = inactiveTable;
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

    }
}
