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
    public partial class Form5 : Form
    {
        ListViewItem[] allitems;
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        public Form5()
        {
            InitializeComponent();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            connection.Open();
            
            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from VoucherLog order by VoucherID ASC", connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
             allitems = new ListViewItem[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[0].ToString());
                OleDbCommand command = new OleDbCommand("select FirstName, LastName from Customers where CustomerID=" + dr[1].ToString(), connection);
                //MessageBox.Show(listitem.Text);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                listitem.SubItems.Add((reader.GetString(0) + " " + reader.GetString(1)));
                reader.Close();
                command.Dispose();
                for (int x = 2; x < 8; x++)
                {
                    listitem.SubItems.Add(dr[x].ToString());
                }
                allitems[i] = listitem;
                listView1.Items.Add(listitem);
            }
            connection.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                DialogResult dialogResult = MessageBox.Show("Selected Voucher will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("select VoucherID, CustomerID, CreditItem from VoucherLog where VoucherID=" + listView1.SelectedItems[0].Text, connection);
                    OleDbDataReader reader = command.ExecuteReader();
                    reader.Read();
                    string log = DateTime.Now.ToString("g") + ": Voucher deletet: id:" + reader.GetInt32(0) + " Customer ID: " + reader.GetInt32(1) + " Credit Item: " + reader.GetString(2) + ";";
                    using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                    {
                        wr.WriteLine(log);
                    }

                    //wr.WriteLine(log);
                    reader.Close();
                    command.CommandText = "delete * from VoucherLog where VoucherID=" + listView1.SelectedItems[0].Text;
                    command.ExecuteNonQuery();
                    command.Dispose();
                    connection.Close();
                    listView1.Items.Clear();
                    Form5_Load(null, null);
                }
            }
            else MessageBox.Show("select a voucher");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form6 f = new Form6();
            f.ShowDialog();
            listView1.Items.Clear();
            Form5_Load(null, null);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form7 f;
            if (listView1.SelectedItems.Count>0) f = new Form7(listView1.SelectedItems[0].Text);
            else 
             f = new Form7(listView1.Items[0].Text);
            f.ShowDialog();
            listView1.Items.Clear();
            Form5_Load(null, null);
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            button3_Click(null, null);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            if (textBox1.Text != null && textBox1.Text != "")
            {
                
                foreach (ListViewItem itm in allitems)
                {
                    if (itm.SubItems[1].Text.ToLower().Contains(textBox1.Text.ToLower())) listView1.Items.Add(itm);
                }
            }
            else listView1.Items.AddRange(allitems);
        }
    }
}
