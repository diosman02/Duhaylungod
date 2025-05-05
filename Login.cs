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

    public partial class Login: Form
    {
        Form4 form4 = new Form4();
        public Login()
        {
            InitializeComponent();
        }       

        private void btnLogin_Click_1(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx");
            Worksheet sheet = book.Worksheets[0];
            int row = sheet.Rows.Length + 1;
            bool logs = false;
            for (int i = 2; i <= row; i++)
            {
                if (sheet.Range[i, 7].Value == txtUsername.Text && sheet.Range[i, 8].Value == txtPassword.Text)
                {
                    Mylogs mylogs = new Mylogs();

                    mylogs.insertLogs(txtUsername.Text, "Logged In");
                    form4.lblIdentifier.Text = sheet.Range[i, 1].Value;




                    logs = true;

                    break;


                }
                else
                {
                    logs = false;
                }
            }
            if (logs == true)
            {
                MessageBox.Show("You are success to continue to another form");
                form4.Show();
                this.Hide();

            }
            else
            {
                MessageBox.Show("Invalid!");

            }
        }
    }
}
