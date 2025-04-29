using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;

namespace Duhaylungod
{
    public partial class Form2: Form
    {
        Workbook book = new Workbook();

        public Form2()
        {
            InitializeComponent();
            //LoadExcelFile();
            //showStudent("1");
        }
        public void LoadExcelFile()
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            dtgInfo.DataSource = dt;
        }

        public void insert(string name, string gender, string hobbies, string color, string saying)
        {
            int i = dtgInfo.Rows.Add();
            dtgInfo.Rows[i].Cells[0].Value = name;
            dtgInfo.Rows[i].Cells[1].Value = gender;
            dtgInfo.Rows[i].Cells[2].Value = hobbies;
            dtgInfo.Rows[i].Cells[3].Value = color;
            dtgInfo.Rows[i].Cells[4].Value = saying;
        }
        public void update(int id, string name, string gender, string hobbies, string color, string saying)
        {

            dtgInfo.Rows[id].Cells[0].Value = name;
            dtgInfo.Rows[id].Cells[1].Value = gender;
            dtgInfo.Rows[id].Cells[2].Value = hobbies;
            dtgInfo.Rows[id].Cells[3].Value = color;
            dtgInfo.Rows[id].Cells[4].Value = saying;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult Yes = MessageBox.Show("Are you sure you want to delete the selected info?", "Notice", MessageBoxButtons.YesNo);

            if (Yes == DialogResult.Yes)
            {
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx");
                Worksheet sh = book.Worksheets[0];
                int row = dtgInfo.CurrentCell.RowIndex + 2;

                sh.Range[row, 13].Value = "0";

                book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx", ExcelVersion.Version2016);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            dtgInfo.ClearSelection();



            if (string.IsNullOrEmpty(txtSearch.Text))
            {
                MessageBox.Show("Please type on the search bar!", "Notice!");
            }
            else
            {
                foreach (DataGridViewRow row in dtgInfo.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(txtSearch.Text))
                    {
                        Mylogs mylogs = new Mylogs();

                        row.Selected = true;
                        break;
                    }
                }

            }
        }

        private void dtgInfo_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int r = dtgInfo.CurrentCell.RowIndex;
            Form1 f1 = (Form1)Application.OpenForms["Form1"];
            f1.txtName.Text = dtgInfo.Rows[r].Cells[0].Value.ToString();
            //f1.radMale.Text = dtgInfo.Rows[r].Cells[1].Value.ToString();
            //f1.radFemale.Text = dtgInfo.Rows[r].Cells[1].Value.ToString();
            string gender = dtgInfo.Rows[r].Cells[1].Value.ToString();
            if (gender == "Male")
            {
                f1.radMale.Checked = true;
            }
            else
            {
                f1.radFemale.Checked = true;
            }
            //f1.chkBasketball.Text = dtgInfo.Rows[r].Cells[2].Value.ToString();
            //f1.chkVolleyball.Text = dtgInfo.Rows[r].Cells[2].Value.ToString();
            //f1.chkSoccer.Text = dtgInfo.Rows[r].Cells[2].Value.ToString();
            string hobbies = dtgInfo.Rows[r].Cells[2].Value.ToString();
            if (hobbies == "Basketball")
            {
                f1.chkBasketball.Checked = true;
            }
            if (hobbies == "Volleyball")
            {
                f1.chkVolleyball.Checked = true;
            }
            if (hobbies == "Soccer")
            {
                f1.chkSoccer.Checked = true;
            }
            f1.cmbFavColor.Text = dtgInfo.Rows[r].Cells[3].Value.ToString();
            f1.txtSaying.Text = dtgInfo.Rows[r].Cells[4].Value.ToString();

            f1.btnAdd.Visible = false;
            f1.btnUpdate.Visible = true;
        }

        public void showStudent(string status)
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\BOOK.xlsx");

            Worksheet sh = book.Worksheets[0];
            DataTable dt = sh.ExportDataTable();
            DataRow[] row = dt.Select("Status" + status);

            foreach (DataRow r in row)
            {
                dtgInfo.Rows.Insert((int)r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13]);
            }

        }


    }
}
