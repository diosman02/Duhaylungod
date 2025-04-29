using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Duhaylungod
{
    class Mylogs
    {
        private const string Format = "hh:mm:ss";

        Workbook book = new Workbook();

        public void insertLogs(string user, string message)
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx"); Worksheet sh = book.Worksheets[1];
            int row = sh.Rows.Length + 1;
            sh.Range[row, 1].Value = user;
            sh.Range[row, 2].Value = message;
            sh.Range[row, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sh.Range[row, 4].Value = DateTime.Now.ToString(Format);
            book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx", ExcelVersion.Version2016);

        }
        public void ShowLogs(DataGridView d)
        {

            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx");
            Worksheet sh = book.Worksheets[1];
            DataTable dt = sh.ExportDataTable();
            d.DataSource = dt;


        }
    }
}
