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
using System.IO;

namespace CATERING_INVOICE_and_CREDIT_VOUCHER
{
    public partial class Form17 : Form
    {

        ListViewItem[] allitems;
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");

        public Form17()
        {
            InitializeComponent();
        }

        private void Form17_Load(object sender, EventArgs e)
        {
            if (connection.State != ConnectionState.Open)
                connection.Open();

            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from Payments order by PaymentID ASC", connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            allitems = new ListViewItem[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[0].ToString());
                OleDbCommand command = new OleDbCommand("select FirstName, LastName from Customers where CustomerID=" + dr[2].ToString(), connection);
                //MessageBox.Show(listitem.Text);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                listitem.SubItems.Add((reader.GetString(0) + " " + reader.GetString(1)));
                reader.Close();
                command.Dispose();
                listitem.SubItems.Add(dr[1].ToString());
                listitem.SubItems.Add(dr[3].ToString());
                listitem.SubItems.Add(dr[4].ToString());
                listitem.SubItems.Add(dr[5].ToString());
                listitem.SubItems.Add(dr[6].ToString());
                allitems[i] = listitem;
                listView1.Items.Add(listitem);
            }
            connection.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count>0)
            {
                Form19 f = new Form19(listView1.SelectedItems[0].Text);
                f.ShowDialog();
                listView1.Items.Clear();
                Form17_Load(null, null);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form18 f = new Form18();
            f.ShowDialog();
            listView1.Items.Clear();
            Form17_Load(null, null);
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            button2_Click(null, null);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {


                DialogResult dialogResult = MessageBox.Show("Selected Payment will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("select PaymentID, PaymentAmount, InvoiceID, customerID from Payments where PaymentID=" + listView1.SelectedItems[0].Text, connection);
                    OleDbDataReader reader = command.ExecuteReader();
                    reader.Read();
                    double amount = reader.GetDouble(1);
                    int invid = reader.GetInt32(2);
                    OleDbCommand cmd = new OleDbCommand("select FirstName, LastName from Customers where customerID=" + reader.GetInt32(3), connection);
                    OleDbDataReader rdr = cmd.ExecuteReader();
                    rdr.Read();
                    string log = DateTime.Now.ToString("g") + ": Payment deletet: id:" + reader.GetInt32(0) + " Customer: " + rdr.GetString(0) + " " + rdr.GetString(1) + " Amount: " + reader.GetDouble(1) + " Invoice:" + reader.GetInt32(2) + ";";
                    using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                    {
                        wr.WriteLine(log);
                    }

                    //wr.WriteLine(log);
                    reader.Close();
                    command.CommandText = "delete * from Payments where PaymentID=" + listView1.SelectedItems[0].Text;
                    command.ExecuteNonQuery();
                    command.CommandText = "update InvoiceHeader set Payments=Payments-" + amount.ToString().Replace(',', '.') + ", BalanceDue=BalanceDue+" + amount.ToString().Replace(',', '.')+" where InvoiceID="+invid;
                    command.ExecuteNonQuery();
                    command.Dispose();
                    connection.Close();
                    MessageBox.Show("Payment Deleted");
                    listView1.Items.Clear();
                    Form17_Load(null, null);
                }
            }
            else MessageBox.Show("select a payment");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox1.Text != null)
            {
                if (radioButton1.Checked)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem itm in allitems)
                    {
                        if (itm.SubItems[1].Text.Contains(textBox1.Text)) listView1.Items.Add(itm);
                    }
                }
                else
                {
                    if (radioButton2.Checked)
                    {
                        listView1.Items.Clear();
                        foreach (ListViewItem itm in allitems)
                        {
                            if (itm.SubItems[2].Text.Contains(textBox1.Text)) listView1.Items.Add(itm);
                        }
                    }
                    else MessageBox.Show("select search mode");
                }
            }
            else
            {
                listView1.Items.Clear();
                listView1.Items.AddRange(allitems);
            }
        }
    }
}
