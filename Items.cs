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
            gett();
        }

        void gett()
        {
            textBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection col = new AutoCompleteStringCollection();
            string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\kkk.xls" + ";Extended Properties='Excel 8.0;HDR=Yes;';";
            OleDbConnection conn = new OleDbConnection(con);
            conn.Open();
            OleDbCommand cmd = new OleDbCommand("Select name from [Sheet1$] ", conn);
            OleDbDataReader rd = cmd.ExecuteReader();
            while(rd.Read())
            {
                string s = rd["name"].ToString();
                col.Add(s);

            }
            textBox1.AutoCompleteCustomSource = col;

        }
        private void button1_Click(object sender, EventArgs e)
        {
            string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\kkk.xls" + ";Extended Properties='Excel 8.0;HDR=Yes;';";
            OleDbConnection conn = new OleDbConnection(con);
            OleDbDataAdapter ad = new OleDbDataAdapter("Select * from [Sheet1$] ", conn);
            OleDbCommand cmd = new OleDbCommand("Select name from [Sheet1$] ", conn);
            DataTable dt = new DataTable();
            ad.Fill(dt);
            dataGridView1.DataSource = dt;
            
         }
        
       private void Items_Load(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }
    }
}
