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
    public partial class Form2 : Form
    {
        ListViewItem[] allitems;
        //string connectionstring = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source="+AppDomain.CurrentDomain.BaseDirectory+"umbertos.accdb; Persist Security Info=False;";
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        bool search;
        public Form2(bool sh)
        {
            search = sh;
            InitializeComponent();
        }
        public bool valid { get; set; }
        private void Form2_Load(object sender, EventArgs e)
        {
            if (search) button3.Text = "Select";
            connection.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from Customers order by CustomerID asc", connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            allitems = new ListViewItem[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[0].ToString());
                for (int x=1; x<22; x++)
                {
                    listitem.SubItems.Add(dr[x].ToString());
                }
                allitems[i] = listitem;
                listView1.Items.Add(listitem);
            }
            connection.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form3 f = new Form3();
            f.ShowDialog();
            listView1.Items.Clear();
            Form2_Load(null, null);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                DialogResult dialogResult = MessageBox.Show("Selected customer will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("select FirstName, LastName from Customers where CustomerID=" + listView1.SelectedItems[0].Text, connection);
                    OleDbDataReader reader = command.ExecuteReader();
                    reader.Read();
                    string log = DateTime.Now.ToString("g") + ": customer deleted id:" + listView1.SelectedItems[0].Text + " First Name: " + reader.GetString(0) + " Last Name: " + reader.GetString(1) + ";";
                    using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                    {
                        wr.WriteLine(log);
                    }

                    //wr.WriteLine(log);
                    reader.Close();
                    command.CommandText = "delete * from Customers where CustomerID=" + listView1.SelectedItems[0].Text;
                    command.ExecuteNonQuery();
                    command.CommandText = "delete * from Voucherlog where CustomerID=" + listView1.SelectedItems[0].Text;
                    command.ExecuteNonQuery();
                    command.CommandText = "select InvoiceId from InvoiceHeader where CustomerID=" + listView1.SelectedItems[0].Text;
                    OleDbDataReader rdr = command.ExecuteReader();
                    List<int> idlist = new List<int>();
                    while (rdr.Read())
                    {
                        idlist.Add(rdr.GetInt32(0));
                    }
                    rdr.Close();
                    foreach (int i in idlist)
                    {
                        command.CommandText = "delete * from InvoiceDetails where invoiceID=" + i;
                        command.ExecuteNonQuery();
                    }
                    command.CommandText = "delete * from InvoiceHeader where CustomerID=" + listView1.SelectedItems[0].Text;
                    command.ExecuteNonQuery();
                    command.CommandText = "delete * from Payments where CustomerID=" + listView1.SelectedItems[0].Text;
                    command.ExecuteNonQuery();
                    command.Dispose();
                    connection.Close();
                    listView1.Items.Clear();
                    Form2_Load(null, null);
                }
            }
            else MessageBox.Show("Select a cutomer.");
        }
        public string name { get; set; }
        public int custid { get; set; }
        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                if (!search)
                {


                    Form4 f;
                    if (listView1.SelectedItems.Count == 0) f = new Form4(listView1.Items[0].Text);
                    else
                        f = new Form4(listView1.SelectedItems[0].Text);
                    f.ShowDialog();
                    listView1.Items.Clear();
                    Form2_Load(null, null);
                }
                else
                {
                    name = listView1.SelectedItems[0].SubItems[2].Text + " " + listView1.SelectedItems[0].SubItems[2].Text;
                    custid=int.Parse(listView1.SelectedItems[0].Text);
                    valid = true;
                    this.Close();
                }
            }
        }

        

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            if (textBox1.Text != null && textBox1.Text != "")
            {
                foreach (ListViewItem itm in allitems)
                {
                    if (itm.SubItems[2].Text.ToLower().Contains(textBox1.Text.ToLower()) || itm.SubItems[3].Text.ToLower().Contains(textBox1.Text.ToLower()) || itm.SubItems[8].Text.ToLower().Contains(textBox1.Text.ToLower())) listView1.Items.Add(itm);
                }
            }
            else listView1.Items.AddRange(allitems);
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            button3_Click(null, null);
        }
    }
}
