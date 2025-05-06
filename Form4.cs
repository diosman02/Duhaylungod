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
    public partial class Form4: Form
    {
        Workbook book = new Workbook();
        Form2 f2 = new Form2();
        public Form4()
        {
            InitializeComponent();

            //Hobbies
            lblCountBasketball.Text = showCount(4, "Basketball").ToString();
            lblCountVolleyball.Text = showCount(4, "Basketball").ToString();

            //Color
            lblCountRed.Text = showCount(10, "Basketball").ToString();
            lblCountBlue.Text = showCount(10, "Basketball").ToString();

            //Course
            lblCountBSIT.Text = showCount(12, "Basketball").ToString();
            lblCountBEED.Text = showCount(12, "Basketball").ToString();

            //Gender
            lblCountMale.Text = showCount(2, "Basketball").ToString();
            lblCountFemale.Text = showCount(2, "Basketball").ToString();

            // Active/Inactive
            lblActive.Text = showCount(13, "1").ToString(); // Count active entries
            lblInactive.Text = showCount(13, "0").ToString();

        }

        private void btnNewStudent_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.Show();
        }

        public int showCount(int c, string val)
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx");
            Worksheet sh = book.Worksheets[0];

            int row = sh.Rows.Length;
            int counter = 0;

            for (int i = 2; i <= row; i++)
            {
                if (sh.Range[i, c].Value.Trim() == val.Trim()) ;
                {
                    counter++;
                }
            }

            return counter;

        }

        private void btnActive(object sender, EventArgs e)
        {
            Active activeForm = new Active();
            activeForm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Inactive inactiveForm = new Inactive();
            inactiveForm.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Logs logsform = new Logs();
            logsform.Show();
        }
    }
}
