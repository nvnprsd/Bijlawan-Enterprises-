using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;


namespace BE
{
    public partial class Genratebill : Form
    {
       
        public Genratebill()
        {
            InitializeComponent();
           
        }
        BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);

        private void button1_Click(object sender, EventArgs e)
        {
            Document dt = new Document(iTextSharp.text.PageSize.A4, 10, 10, 10, 10);
            PdfWriter wr=  PdfWriter.GetInstance(dt, new FileStream("D:/Projects/Convert ScreenToPDF/hello.pdf", FileMode.Create));
            dt.Open();
             BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
            iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font t8 = new iTextSharp.text.Font(bfTimes, 11, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

            iTextSharp.text.Font t3, t1 = new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font t9 = new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.UNDERLINE, BaseColor.BLACK);
            t3 = new iTextSharp.text.Font(bfTimes, 18, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            Paragraph pg =new Paragraph("                                                                                                      Tax Invoice                                                                                                      "+Environment.NewLine, t9);                
            pg.Alignment = Element.ALIGN_CENTER;
            Paragraph p3 = new Paragraph("GSTIN : 09CYDPD5055M1Z6                                                                                                                                            Mob :9876543210", t1);
            Paragraph p2 = new Paragraph("BIJALWAN ENTERPRISES", t3);
            p2.Alignment = Element.ALIGN_CENTER;
            Paragraph u3 = new Paragraph("105,SF CHI-04,Greater Noida (Gautam Budha Nagar)"+Environment.NewLine+"Cleaning Meterial, Stationery, Disposable Bags and Gloves", t8);
            u3.Alignment = Element.ALIGN_CENTER;
            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 90;
            PdfPCell cell = new PdfPCell();
            cell.HorizontalAlignment = 1;
            float[] w = {60f,40f };
            table.AddCell(new Phrase("M/s-----------------------------------------------------------------------",t2));
            table.AddCell(new Phrase("Invoice No. " + Environment.NewLine + "Invoice Date.------------" + Environment.NewLine + "Transportation Mode.-----------------" + Environment.NewLine + "Vehicle Number.------------" + Environment.NewLine + "Date of Supply.------------",t2));
            table.SetWidths(w);
            table.AddCell(cell); //  cell.Colspan = 3;
            table.SpacingBefore = 10;
            table.SpacingAfter = table.CalculateHeights();
             iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance("D:/helll.png");
            img.ScaleToFit(120,140);
            img.Alignment = Element.ALIGN_RIGHT;
            PdfPTable tb = new PdfPTable(6);
            tb.WidthPercentage = 90;
            tb.AddCell(new Phrase("S.No",t2));
            float[] wi = { 5f, 45f, 10f, 7f, 13f, 20f };

            tb.AddCell(new Phrase("Description",t2));
            tb.AddCell(new Phrase("HSN Code",t2));
            tb.AddCell(new Phrase("Qty",t2));
            tb.AddCell(new Phrase("Rate",t2));
            tb.AddCell(new Phrase("Amount",t2));
            tb.SetWidths(wi);
            tb.SpacingAfter = tb.CalculateHeights();
            dt.Add(pg); dt.Add(p3);
            dt.Add(p2);dt.Add(u3);
            //memo and ttls
            dt.Add(table);dt.Add(tb);
            //repeated
            dt.Add(column()); dt.Add(column());
            dt.Add(netamt()); dt.Add(gstamt());
            dt.Add(finale());
            dt.Add(img);
            dt.Close();

        }
        PdfPTable column()
        {

            iTextSharp.text.Font t1 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLDITALIC, BaseColor.BLACK);
            PdfPTable tb = new PdfPTable(6);
            tb.WidthPercentage = 90;
            tb.AddCell("1");
            float[] wi = { 5f, 45f, 10f, 7f, 13f, 20f };

            tb.AddCell(new Phrase("rate dfg d gd f gf",t1));
            tb.AddCell(new Phrase("ABCDEF", t1));
            tb.AddCell(new Phrase("112", t1));
            tb.AddCell(new Phrase("110000", t1));
            tb.AddCell(new Phrase("990000", t1));
            tb.SetWidths(wi);
            tb.SpacingAfter = tb.CalculateHeights();
            return tb;
        }
        PdfPTable netamt()
        {
            iTextSharp.text.Font t1 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLDITALIC, BaseColor.BLACK);

            PdfPTable net = new PdfPTable(3);
            net.WidthPercentage = 90;
            float[] f = { 60f,20f, 20f };
            net.AddCell(" ");net.AddCell(new Phrase("Net Amount", t1));
            net.AddCell(new Phrase("1200", t2));
            net.SetWidths(f);
            net.SpacingAfter = net.CalculateHeights();

            return net;
        }
        PdfPTable gstamt()
        {
            iTextSharp.text.Font t1 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLDITALIC, BaseColor.BLACK);

            PdfPTable gst = new PdfPTable(2);
            gst.WidthPercentage = 90;
            gst.AddCell(new Phrase("         BANK A/C DETAILS." + Environment.NewLine + "  Bank Name : Bank Of Baroda" + Environment.NewLine + "   A/C No. 56200200000135" + Environment.NewLine + " IFSC Code: BARBOKASNAX", t2));
            PdfPTable nst = new PdfPTable(2);
            nst.AddCell(new Phrase("Discount %", t1));
            nst.AddCell(new Phrase("00000", t2));
            nst.AddCell(new Phrase("Amt After Discount", t1));
            nst.AddCell(new Phrase("00000", t2));
            nst.AddCell(new Phrase("ADD: CGST @", t1));
            nst.AddCell(new Phrase("00000",t2));
            PdfPCell pd = new PdfPCell(nst);
            gst.AddCell(pd);
            gst.SpacingAfter = gst.CalculateHeights();
            float[] f = { 60f, 40f };
            gst.SetWidths(f);
            return gst;
        }
        PdfPTable finale ()
        {
            iTextSharp.text.Font t1 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLDITALIC, BaseColor.BLACK);

            PdfPTable final = new PdfPTable(2);
            final.WidthPercentage = 90;
            float[] f = {60f,40f };
            PdfPTable ns = new PdfPTable(1);
            ns.AddCell(new Phrase("Amt in Words..............................................................................................................................................................................................................",t2));
            ns.AddCell(new Phrase("This is Computer Generated Recipt Does Not Required Any Physical Signature.", new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLUE)));
            PdfPCell cc = new PdfPCell(ns);
            
            final.AddCell(cc);
            PdfPTable nst = new PdfPTable(2);
            nst.AddCell(new Phrase("ADD: SGST @", t1));
            nst.AddCell(new Phrase("00000", t2));
            nst.AddCell(new Phrase("ADD: IGST @", t1));
            nst.AddCell(new Phrase("00000", t2));
            nst.AddCell(new Phrase("Other Charges", t1));
            nst.AddCell(new Phrase("00000", t2));
            nst.AddCell(new Phrase("Total Amount", t1));
            nst.AddCell(new Phrase("100000000", t2));
            PdfPCell c = new PdfPCell(nst);
            final.SetWidths(f);
            final.AddCell(c);
            return final;


        }



    }
}
