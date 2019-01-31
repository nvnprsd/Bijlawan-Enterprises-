using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Data.OleDb;
using System.Configuration;

namespace BE
{
    public partial class Genratebill : Form
    {

        public Genratebill()
        {
            InitializeComponent();
           gett();
            it();

        }
        void it()
        {
            textBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox2.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection col = new AutoCompleteStringCollection();
            string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\databasefile.xls" + ";Extended Properties='Excel 8.0;HDR=Yes;';";
            OleDbConnection conn = new OleDbConnection(con);
            conn.Open();
            OleDbCommand cmd = new OleDbCommand("Select name from [Sheet1$] ", conn);
            OleDbDataReader rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string s = rd["name"].ToString();
                col.Add(s);

            }
            textBox2.AutoCompleteCustomSource = col;

        }
        void gett()
        {
            textBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection col = new AutoCompleteStringCollection();
            string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\databasefile.xls" + ";Extended Properties='Excel 8.0;HDR=Yes;';";
            OleDbConnection conn = new OleDbConnection(con);
            conn.Open();
            OleDbCommand cmd = new OleDbCommand("Select GSTIN from [Sheet2$] ", conn);
            OleDbDataReader rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string s = rd["GSTIN"].ToString();
                col.Add(s);//

            }
            textBox1.AutoCompleteCustomSource = col;

        }
        BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string path = @"D:\Bijalwan Enterprise\Transport\";
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                GETpdf("Transport Copy", path);

                path = @"D:\Bijalwan Enterprise\Duplicate\";
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                GETpdf("Duplicate Copy", path);

                path = @"D:\Bijalwan Enterprise\Original\";
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                    var di =new  DirectoryInfo(path);
                    di.Attributes = FileAttributes.Hidden;
                    di.Attributes |= FileAttributes.Encrypted;
                    di.Attributes |= FileAttributes.NotContentIndexed;

                }
                GETpdf("Original Copy", path);
                var d = new DirectoryInfo(@"d:\Bijalwan Enterprise");
                d.Attributes = FileAttributes.Hidden;
                MessageBox.Show("Bill Generated  " + DateTime.Now.ToShortDateString() + invoice.Text +textBox1.Text+ ".pdf");
                this.Hide();
                System.Diagnostics.Process.Start("explorer.exe", @"D:\Bijalwan Enterprise");
            }
            catch(Exception ex) {
                MessageBox.Show("Error in Creation Directory" +Environment.NewLine+"Error code32420"+ex.ToString(),MessageBoxIcon.Error.ToString());

            }
            

        }
        private string numtoword(int number)
        {

            string words = "";
            if ((number / 1000000) > 0)
            {
                words += numtoword(number / 100000) + " LAKH";
                number %= 1000000;
            }
            if ((number / 1000) > 0)
            {
                words += numtoword(number / 1000) + " THOUSAND ";
                number %= 1000;
            }
            if ((number / 100) > 0)
            {
                words += numtoword(number / 100) + " HUNDRED ";
                number %= 100;
            }
           if (number > 0)
            {
                if (words != "") words += "AND ";
                var unitsMap = new[]
                {
            "ZERO", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"
        };
                var tensMap = new[]
                {
            "ZERO", "TEN", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"
        };
                if (number < 20) words += unitsMap[number];
                else
                {
                    words += tensMap[number / 10];
                    if ((number % 10) > 0) words += " " + unitsMap[number % 10];
                }
            }
            return words;
           
        }
        private void GETpdf(string a,string path)
        {
            try
            {
                Document dt = new Document(iTextSharp.text.PageSize.A4.Rotate(), 10, 10, 20, 10);
                PdfWriter wr = PdfWriter.GetInstance(dt, new FileStream(path+a.Substring(0,2) + DateTime.Now.ToShortDateString() + invoice.Text + textBox1.Text + ".pdf", FileMode.Create));
                dt.Open();
                BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
                iTextSharp.text.Font tt = new iTextSharp.text.Font(BaseFont.CreateFont(BaseFont.TIMES_ITALIC, BaseFont.CP1252, false), 11, iTextSharp.text.Font.ITALIC, BaseColor.BLACK);
                iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                iTextSharp.text.Font t8 = new iTextSharp.text.Font(bfTimes, 11, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                iTextSharp.text.Font t3, t1 = new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                iTextSharp.text.Font t9 = new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.UNDERLINE, BaseColor.BLACK);
                t3 = new iTextSharp.text.Font(bfTimes, 18, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
                Paragraph pg = new Paragraph("                                                                                                                                               Tax Invoice                                                                                                                 * " + a + "             " + Environment.NewLine, t9);
                pg.Alignment = Element.ALIGN_CENTER;
                Paragraph p3 = new Paragraph("            GSTIN : 09CYDPD5055M1Z6                                                                                                                                                                                                      Mob :9560078377,8882414574", t1);
                Paragraph p2 = new Paragraph("BIJALWAN ENTERPRISES", t3);
                p2.Alignment = Element.ALIGN_CENTER;
                Paragraph u3 = new Paragraph("105,SF CHI-04,Greater Noida (Gautam Buddha Nagar)" + Environment.NewLine + "Cleaning Meterial, Stationery, Disposable Bags and Gloves", t8);
                u3.Alignment = Element.ALIGN_CENTER;
                PdfPTable table = new PdfPTable(3);
                table.WidthPercentage = 90;
                float[] w = { 35f, 35f, 30f };

                table.AddCell(new Phrase("Consignee:" + Environment.NewLine + " M/s " + cmpnyname.Text + Environment.NewLine + address.Text + Environment.NewLine + contact.Text + "", t2));
                table.AddCell(new Phrase("Shipping Address" + Environment.NewLine + " M/s " + scmpny.Text + Environment.NewLine + Saddress.Text + Environment.NewLine + Scontact.Text + "", t2));
                table.AddCell(new Phrase("Invoice No. " + invoice.Text + Environment.NewLine + "Invoice Date. " + DateTime.Now.ToShortDateString() + "" + Environment.NewLine + "Transportation Mode." + Tmode.Text + "" + Environment.NewLine + "Vehicle Number." + Vnum.Text + "" + Environment.NewLine + "Date of Supply." + dos.Text + "", t2));
                table.SetWidths(w);
                //  cell.Colspan = 3;
                table.SpacingBefore = 10;
                table.SpacingAfter = table.CalculateHeights();
                //iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance("D:/helll.png");
                //img.ScaleToFit(120, 140);
                //img.Alignment = Element.ALIGN_RIGHT;
                PdfPTable tb = new PdfPTable(12);
                tb.WidthPercentage = 90;
                tb.AddCell(new Phrase("S.No", t2));
                float[] wi = { 3f, 27f, 8f, 5f, 7f, 7f, 8f, 6f, 6f, 6f, 7f, 10f };

                tb.AddCell(new Phrase("Product Name", t2));
                tb.AddCell(new Phrase("HSN Code", t2));
                tb.AddCell(new Phrase("Qty", t2));
                tb.AddCell(new Phrase("Rate", t2));
                tb.AddCell(new Phrase("Unit Discount" + Environment.NewLine + "   (%)", t2));
                tb.AddCell(new Phrase("Amt after Discount", t2));
                tb.AddCell(new Phrase("CGST" + Environment.NewLine + "   (%)", t2));
                tb.AddCell(new Phrase("SGST" + Environment.NewLine + "   (%)", t2));
                tb.AddCell(new Phrase("IGST" + Environment.NewLine + "   (%)", t2));
                tb.AddCell(new Phrase("Total Tax Amt", t2));
                tb.AddCell(new Phrase("Final Amount", t2));


                tb.SetWidths(wi);
                tb.SpacingAfter = tb.CalculateHeights();
                dt.Add(pg); dt.Add(p3);
                dt.Add(p2); dt.Add(u3);
                //memo and ttls
                dt.Add(table); dt.Add(tb);
                //repeated
                int s = 0;
                double Tcgst = 0, Tsgst = 0, Tigst = 0, Final = 0, Ttaxamt = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    s++;
                    string name = row.Cells[0].Value.ToString();
                    string hsn = row.Cells[1].Value.ToString(); double qty = Convert.ToDouble(row.Cells[2].Value.ToString());
                    double rate = Convert.ToDouble(row.Cells[3].Value.ToString()); double dis = Convert.ToDouble(row.Cells[4].Value.ToString());
                    double amtd = rate - ((rate * dis) / 100); double cgst = Convert.ToDouble(row.Cells[5].Value.ToString());
                    double sgst = Convert.ToDouble(row.Cells[6].Value.ToString()); double igst = Convert.ToDouble(row.Cells[7].Value.ToString());
                    double ttax =Math.Round( qty * (amtd * (cgst + sgst + igst) / 100),2); double amt =Math.Round( (ttax + (amtd * qty)),2);
                    dt.Add(column(s, name, hsn, qty, rate, dis, amtd, cgst, sgst, igst, ttax, amt));
                    Tcgst += (qty * (rate * cgst) / 100);
                    Tsgst += (qty * (rate * sgst) / 100);
                    Tigst += (qty * (rate * igst) / 100);
                    Final += amt; Ttaxamt += ttax;
                }
                dt.Add(other());
                Tcgst = Math.Round(Tcgst, 2);
                Tsgst = Math.Round(Tsgst, 2);
                Tigst = Math.Round(Tigst, 2);
                Final = Math.Round(Final + Convert.ToDouble(othercharges.Text), 2);
                dt.Add(netamt(Tcgst, Tsgst, Tigst, Ttaxamt, Final));
                dt.Add(finale(Final));
                dt.Add(gstamt());
                Paragraph ad = new Paragraph("Contact for These type of Billing Softwares and other Software & Websites. naveenprasad364@gmail.com or 7060629794", tt);
                 ad.Alignment = Element.ALIGN_CENTER;
                ad.SpacingBefore = 40f;
                dt.Add(ad);
                dt.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show("unable to Print Bill System File Missing" + Environment.NewLine + "Error code 12420" + Environment.NewLine +ex.ToString(), MessageBoxIcon.Error.ToString());

            }
        }
        PdfPTable column(int s,string name, string hsn, double qty, double rate, double discount, double amtdis, double cgst, double sgst, double igst, double ttax, double amt)
        {

            iTextSharp.text.Font t1 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLDITALIC, BaseColor.BLACK);
            PdfPTable tb = new PdfPTable(12);
            tb.WidthPercentage = 90;
            tb.AddCell(s.ToString());
            float[] wi = { 3f, 27f, 8f, 5f, 7f, 7f, 8f, 6f, 6f, 6f, 7f, 10f };

            tb.AddCell(new Phrase(name, t1));
            tb.AddCell(new Phrase(hsn, t1));
            tb.AddCell(new Phrase(qty.ToString(), t1));
            tb.AddCell(new Phrase(rate.ToString(), t1));
            tb.AddCell(new Phrase(discount.ToString()+"%", t1));
            tb.AddCell(new Phrase(amtdis.ToString(), t1));
            tb.AddCell(new Phrase(cgst.ToString() + "%", t1));
            tb.AddCell(new Phrase(sgst.ToString() + "%", t1));
            tb.AddCell(new Phrase(igst.ToString() + "%", t1));
            tb.AddCell(new Phrase(ttax.ToString(), t1));
            tb.AddCell(new Phrase(amt.ToString(), t1));
            tb.SetWidths(wi);
            tb.SpacingAfter = tb.CalculateHeights();
            return tb;
        }
        PdfPTable netamt(double a,double b,double c,double d,double e)
        {
            iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLDITALIC, BaseColor.BLACK);
            double k = e;
            e = Math.Round(e);
            PdfPTable net = new PdfPTable(6);
            net.WidthPercentage = 90;
            float[] f = { 65f, 6f, 6f, 6f, 7f, 10f };
            net.AddCell(new Phrase("Grand Total                       Rounded off ("+k+")", t2));
            net.AddCell(new Phrase(a.ToString(), t2));
            net.AddCell(new Phrase(b.ToString(), t2));
            net.AddCell(new Phrase(c.ToString(), t2));
            net.AddCell(new Phrase(d.ToString(), t2));
            net.AddCell(new Phrase(e.ToString(), t2));
            net.SetWidths(f);
            net.SpacingAfter = net.CalculateHeights();

            return net;
        }
        PdfPTable other()
        {
            iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLDITALIC, BaseColor.BLACK);
            iTextSharp.text.Font t1 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 90;
            float[] w = { 83f, 17f };

            table.AddCell(new Phrase("Other Chareges*", t1));
            table.AddCell(new Phrase(othercharges.Text, t2));
            table.SetWidths(w);
            table.SpacingAfter = table.CalculateHeights();
            return table;
        }
        PdfPTable gstamt()
        {
            iTextSharp.text.Font t1 = t1 = new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLDITALIC, BaseColor.BLACK);

            PdfPTable gst = new PdfPTable(2);
            gst.WidthPercentage = 90;
            gst.AddCell(new Phrase("Terms & Conditions."+Environment.NewLine+"1. Goods once sold will not be taken back." + Environment.NewLine +"2.Interest @ 24 P.A. will be Charged if the payment is not made within the stipulated time." + Environment.NewLine +"3. Subject to 'Gautam Buddha Nagar' Jurisdiction only.", t1));
            PdfPTable tb = new PdfPTable(1);
            tb.AddCell(new Phrase("Reciver's Signature"+Environment.NewLine+" ",t2));
            tb.AddCell(new Phrase("                                       For Bijalwan Enterprises" + Environment.NewLine + Environment.NewLine + Environment.NewLine + "                                             Authorised Signatory", t2));
           
            PdfPCell c = new PdfPCell(tb);
            gst.AddCell(c);
            gst.SpacingAfter = gst.CalculateHeights();
            float[] f = { 60f, 40f };
            gst.SetWidths(f);
            return gst;
        }
        PdfPTable finale(double a)
        {
            iTextSharp.text.Font t1 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            float[] x={60f,40f };
            PdfPTable final = new PdfPTable(2);
            final.WidthPercentage = 90;
            final.AddCell(new Phrase("Amount in Words to Be Collected."+numtoword(Convert.ToInt32( a))+" Rupees Only ", t2));
            final.AddCell(new Phrase("     BANK A/C DETAILS." + Environment.NewLine + "  Bank Name : Bank Of Baroda" + Environment.NewLine + "   A/C No. 56200200000135" + Environment.NewLine + " IFSC Code: BARBOKASNAX", t2));
            // final.AddCell(new Phrase("This is Computer Generated Recipt Does Not Required Any Physical Signature.", new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLUE)));
            final.SetWidths(x);
            return final;


        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }
      
        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        ////private void button4_Click(object sender, EventArgs e)
        ////{       ////    SqlConnection con = new SqlConnection("Server=localhost;Integrated security=SSPI;database=master");
        ////    string s = "CREATE DATABASE MyDatabase ON PRIMARY " +
        ////"(NAME = MyDatabase_Data, " +
        ////"FILENAME = 'C:\\MyDatabaseData.mdf') " +
        ////"LOG ON (NAME = MyDatabase_Log, " +
        ////"FILENAME = 'C:\\MyDatabaseLog.ldf')";
        ////    SqlCommand cmd = new SqlCommand(s, con);
        ////    con.Open();
        ////    cmd.ExecuteNonQuery();
        ////    MessageBox.Show("db created");
        ////    con.Close();
        ////}

        private void button5_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection("Data Source = localhost; Integrated Security = SSPI; " +
                                            "Initial Catalog=MyDatabaseData"))
            {
                SqlCommand cmd = new SqlCommand("create table emp (name varchar(10),mob int);", con);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("table created")
;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            try { string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\databasefile.xls" + ";Extended Properties='Excel 8.0;HDR=Yes;';";
                OleDbConnection conn = new OleDbConnection(con);
                conn.Open();
                OleDbCommand cmd = new OleDbCommand("Select name,address,contact from [Sheet2$] where GSTIN=" + textBox1.Text + "", conn);
                OleDbDataReader rd = cmd.ExecuteReader();
                while (rd.Read())
                {

                    cmpnyname.Text = rd["name"].ToString();
                    address.Text = rd["address"].ToString();
                    contact.Text = rd["contact"].ToString();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Unreachable  excel file or Not Found" +ex.ToString());
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {

            chaku();
        }

       private void chaku()
        {
            try
            {
                string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\\kkk.xls" + ";Extended Properties='Excel 8.0;HDR=Yes;';";
                OleDbConnection conn = new OleDbConnection(con);
                OleDbDataAdapter ad = new OleDbDataAdapter("Select * from [Sheet1$] where name='" + textBox2.Text + "'", conn);
                DataTable d = new DataTable();
                ad.Fill(d);
                foreach (DataGridViewRow rw in dataGridView1.Rows)
                {
                    DataRow dr = d.NewRow();
                    dr[0] = rw.Cells[0].Value; dr[3] = rw.Cells[3].Value; dr[6] = rw.Cells[6].Value; //dr[9] = rw.Cells[9];
                    dr[1] = rw.Cells[1].Value; dr[4] = rw.Cells[4].Value; dr[7] = rw.Cells[7].Value;// dr[10] = rw.Cells[10];
                    dr[2] = rw.Cells[2].Value; dr[5] = rw.Cells[5].Value;// dr[8] = rw.Cells[8]; dr[11] = rw.Cells[11];
                    d.Rows.Add(dr);
                }
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dataGridView1.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold);

                dataGridView1.DataSource = d;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Source Not found" + ex.ToString());
            }
        }
        private void Genratebill_Load(object sender, EventArgs e)
        {
            dos.Text = DateTime.Now.ToShortDateString();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked==true)
            {
                scmpny.Text = cmpnyname.Text;
                Saddress.Text = address.Text;
                Scontact.Text = contact.Text;
            }
            else
            {
                scmpny.Text = "";
                Saddress.Text ="";
                Scontact.Text = "";
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            panel2.Visible = true;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(Convert.ToInt32(e.KeyChar)==13)
            {
                chaku();
                textBox2.Text = "";
            }
        }

        private void contact_Leave(object sender, EventArgs e)
        {
            checkBox1.Focus();
        }

        private void checkBox1_Leave(object sender, EventArgs e)
        {
            scmpny.Focus();
        }

        private void dos_Leave(object sender, EventArgs e)
        {
            button5.Focus();
        }

        private void label17_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}

