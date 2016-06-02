using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CATERING_INVOICE_and_CREDIT_VOUCHER
{
    public partial class Form22 : Form
    {
        OleDbConnection connection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        public Form22()
        {
            InitializeComponent();
        }


        void getpayments()
        {
            listView1.Items.Clear();
            List<ListViewItem> items = new List<ListViewItem>();
            double total = 0;
            int invoices = 0;
            if (connection.State != ConnectionState.Open) connection.Open();
            OleDbCommand command = new OleDbCommand("select invoiceDate, Subtotal, Discount, invoiceID from InvoiceHeader order by InvoiceID asc", connection);
            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                if (dateTimePicker1.Value <= DateTime.Parse(reader.GetString(0)) && dateTimePicker2.Value >= DateTime.Parse(reader.GetString(0)))
                {
                    total = total + reader.GetDouble(1) - ((reader.GetDouble(2) / 100) * reader.GetDouble(1));
                    invoices++;


                    OleDbCommand cmd = new OleDbCommand("select ProductID, DishName, UnitPrice, Quantity from InvoiceDetails where InvoiceID=" + reader.GetInt32(3)+ " order by ProductID asc", connection);
                    OleDbDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        bool exist = false;
                        foreach (ListViewItem itm in items)
                        {
                            if (itm.Text == rdr.GetInt32(0).ToString())
                            {
                                itm.SubItems[3].Text = (double.Parse(itm.SubItems[3].Text) + rdr.GetDouble(3)).ToString();
                                itm.SubItems[4].Text = (double.Parse(itm.SubItems[3].Text) * rdr.GetDouble(2)).ToString();
                                exist = true;
                            }
                        }
                           if(!exist)
                            {
                                ListViewItem itm2 = new ListViewItem(rdr.GetInt32(0).ToString());
                                itm2.SubItems.Add(rdr.GetString(1));
                                itm2.SubItems.Add(rdr.GetDouble(2).ToString());
                                itm2.SubItems.Add(rdr.GetDouble(3).ToString());
                                itm2.SubItems.Add((rdr.GetDouble(2) * rdr.GetDouble(3)).ToString());
                                items.Add(itm2);
                            }
                        }
                    }


                
            }
            label3.Text = "Incoming Payments: $" + total;
            label4.Text = "Total Invoices: " + invoices;
            items = items.OrderBy(itm => int.Parse(itm.SubItems[0].Text)).ToList();
            foreach (ListViewItem itm in items)
            {
                listView1.Items.Add(itm);
            }
        }

        

        private void Form22_Load(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            getpayments();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            getpayments();
        }
    }
}
