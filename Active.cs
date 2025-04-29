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
    public partial class Active: Form
    {
        public Active()
        {
            InitializeComponent();
            LoadActiveData();
        }
        private void LoadActiveData()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();

            // Filter Active Entries
            DataRow[] activeRows = dt.Select("Status = '1'");
            DataTable activeTable = activeRows.CopyToDataTable();
            dataGridView1.DataSource = activeTable; // Assuming dataGridView1 is the name of your DataGridView
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

       
    }
}
