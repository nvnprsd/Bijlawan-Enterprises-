using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace BE
{
    public partial class Items : Form
    {
        public Items()
        {
            InitializeComponent();
            data.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(data.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold);

        }


        private void button1_Click(object sender, EventArgs e)
        {
           

        }

        private void Items_Load(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Genratebill g = new Genratebill();
            g = (Genratebill)Application.OpenForms["Genratebill"];
           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Genratebill g = new Genratebill();
            g.Show();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", @"D:\Bijalwan Enterprise");

        }

        private void button6_Click(object sender, EventArgs e)
        {
            merchant.Visible = false;
            panel1.Visible = false;
            product.Visible = true;
            //string path = @"D:\softDB\";
            //if (!System.IO.Directory.Exists(path))
            //{
            //    System.IO.Directory.CreateDirectory(path);
            //    Microsoft.Office.Interop.Excel.Application oXl;
            //    Microsoft.Office.Interop.Excel._Workbook owb;
            //    Microsoft.Office.Interop.Excel._Worksheet osl;
            //    object val = System.Reflection.Missing.Value;
            //    oXl = new Microsoft.Office.Interop.Excel.Application();
            //    owb = (Microsoft.Office.Interop.Excel._Workbook)(oXl.Workbooks.Add());
            //    osl = (Microsoft.Office.Interop.Excel._Worksheet)owb.ActiveSheet;
            //    osl.Cells[1, 1] = "name";
            //    osl.Cells[1, 2] = "hsn code";
            //    osl.Cells[1, 3] = "qty rate";
            //    osl.Cells[1, 4] = "discount";
            //    osl.Cells[1, 5] = "cgst";
            //    osl.Cells[1, 6] = "sgst";
            //    osl.Cells[1, 7] = "igst";
            //    osl = osl.Next;
            //    osl.Cells[1, 1] = "GSTIN";
            //    osl.Cells[1, 2] = "name";
            //    osl.Cells[1, 3] = "address";
            //    osl.Cells[1, 4] = "email";
            //    osl.Cells[1, 5] = "contact";

            //    oXl.UserControl = false;
            //    owb.SaveAs("D:\\SoftDB\\databasefile.xls");//, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);//, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //    owb.Close();
            //}
            //else
          
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
           
        }

      

        private void button5_Click(object sender, EventArgs e)
        {

            try
            {
                string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\databasefile.xls" + ";Extended Properties=Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(con);
                conn.Open();
                OleDbCommand cmd = new OleDbCommand("insert into [Sheet2$] (GSTIN,name,address,email,contact)values('" + gstin.Text + "','" + name.Text + "','" + address.Text + "','" + email.Text + "','" + contact.Text + "')", conn);
                cmd.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Congsinee Added", MessageBoxButtons.OK.ToString());

            }
            catch
            {
                MessageBox.Show("Please paste the database file into D drive with 'databasefile.xls' name and appropriate format ");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\databasefile.xls" + ";Extended Properties=Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(con);
                conn.Open();
                OleDbCommand cmd = new OleDbCommand("insert into [Sheet1$] (name,hsn,qty,rate,discount,cgst,sgst,igst)values('"+Pname.Text+"','"+hsn.Text+"','"+qty.Text+"','"+rate.Text+"','"+discount.Text+"','"+cgst.Text+"','"+sgst.Text+"','"+igst.Text+"')", conn);
                cmd.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Product  Added", MessageBoxButtons.OK.ToString());
              }
            catch
            {
                MessageBox.Show("Please paste the database file into D drive with 'databasefile.xls' name and appropriate format ");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Comming Soon in new Patch...");


        }

        private void button3_Click(object sender, EventArgs e)
        {
            product.Visible = false;
            merchant.Visible = false;
            panel1.Visible = true;
            try
            {
                string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\databasefile.xls" + ";Extended Properties=Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(con);
                conn.Open();
                OleDbDataAdapter cmd = new OleDbDataAdapter("select * from [Sheet2$]", conn);
                DataTable dt = new DataTable();
                cmd.Fill(dt);
                data.DataSource = dt;
                conn.Close();
               
            }
            catch
            {
                MessageBox.Show("Please paste the database file into D drive with 'databasefile.xls' name and appropriate format ");
            }
        }
        void btn()
        {
            product.Visible = false;
            merchant.Visible = false;
            panel1.Visible = true;
            try
            {
                string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\databasefile.xls" + ";Extended Properties=Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(con);
                conn.Open();
                OleDbDataAdapter cmd = new OleDbDataAdapter("select * from [Sheet1$]", conn);
                DataTable dt = new DataTable();
                cmd.Fill(dt);
                data.DataSource = dt;
                conn.Close();
               
            }
            catch
            {
                MessageBox.Show("Please paste the database file into D drive with 'databasefile.xls' name and appropriate format ");
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            btn();
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            try
            {
                string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\databasefile.xls" + ";Extended Properties=Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(con);
                conn.Open();
                foreach (DataGridViewRow rw in data.Rows)
                {
                    OleDbCommand cmd = new OleDbCommand("update [Sheet1$] set hsn='" + hsn.Text + "'qty='" + qty.Text + "',rate='" + rate.Text + "',discount='" + discount.Text + "',cgst='" + cgst.Text + "',sgst='" + sgst.Text + "',igst='" + igst.Text + "' where name='" + name.Text + "')", conn);
                    cmd.ExecuteNonQuery();
                    cmd = new OleDbCommand("update [Sheet1$] set hsn=" + rw.Cells[1].Value + " where name='" + name.Text + "')", conn);
                    cmd.ExecuteNonQuery();
                  cmd = new OleDbCommand("update [Sheet1$] set qty=" + rw.Cells[2].Value + " where name='" + name.Text + "')", conn);
                    cmd.ExecuteNonQuery();
                   cmd = new OleDbCommand("update [Sheet1$] set rate='" + rw.Cells[1].Value + "' where name='" + name.Text + "')", conn);
                    cmd.ExecuteNonQuery();
                     cmd = new OleDbCommand("update [Sheet1$] set discount='" + rw.Cells[1].Value + "' where name='" + name.Text + "')", conn);
                    cmd.ExecuteNonQuery();
                   cmd = new OleDbCommand("update [Sheet1$] set cgst='" + rw.Cells[1].Value + "' where name='" + name.Text + "')", conn);
                    cmd.ExecuteNonQuery();
                    cmd = new OleDbCommand("update [Sheet1$] set sgst='" + rw.Cells[1].Value + "' where name='" + name.Text + "')", conn);
                    cmd.ExecuteNonQuery();
                    cmd = new OleDbCommand("update [Sheet1$] set igst='" + rw.Cells[1].Value + "' where name='" + name.Text + "')", conn);
                    cmd.ExecuteNonQuery();

                }
                conn.Close();
                MessageBox.Show("Update Done");
                panel1.Visible = false;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Please paste the database file into D drive with 'databasefile.xls' name and appropriate format "+ex.ToString());
            }
        }
       
private void button2_Click_2(object sender, EventArgs e)
        {
            product.Visible = false;
            merchant.Visible = true;
            panel1.Visible = false;

        }

        private void label17_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}

