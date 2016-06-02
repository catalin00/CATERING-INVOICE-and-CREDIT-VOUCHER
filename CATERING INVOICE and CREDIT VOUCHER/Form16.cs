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
    public partial class Form16 : Form
    {
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        ListViewItem[] allitems;
        List<ListViewItem> selectedfilter = new List<ListViewItem>();
        //List<ListViewItem> showitem = new List<ListViewItem>();
        public Form16()
        {
            
            InitializeComponent();
        }
        void updatelist()
        {
            listView1.Items.Clear();
            List<ListViewItem> showitem = new List<ListViewItem>();
            if (textBox1.Text=="" && textBox1.Text==null)
            {
                showitem = datedlist(selectedfilter);
            }
            else
            {
                foreach(ListViewItem itm in selectedfilter)
                {
                    if (itm.SubItems[1].Text.ToLower().Contains(textBox1.Text.ToLower())) showitem.Add(itm);
                }
                showitem = datedlist(showitem);
            }
            foreach (ListViewItem itm in showitem)
            {
                listView1.Items.Add(itm);
            }
        }
        private void Form16_Load(object sender, EventArgs e)
        {
            
            if (connection.State!=ConnectionState.Open) connection.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from InvoiceHeader", connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            allitems = new ListViewItem[dt.Rows.Count];
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                ListViewItem itm = new ListViewItem(dr[0].ToString());
                OleDbCommand command = new OleDbCommand("select FirstName, LastName from Customers where CustomerID=" + dr[1].ToString(), connection);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                itm.SubItems.Add(reader.GetString(0) + " " + reader.GetString(1));
                itm.SubItems.Add(dr[2].ToString());
                itm.SubItems.Add(dr[4].ToString());
                itm.SubItems.Add(dr[5].ToString());
                itm.SubItems.Add(dr[6].ToString());
                itm.SubItems.Add(dr[7].ToString());
                itm.SubItems.Add(dr[8].ToString());
                itm.SubItems.Add(dr[9].ToString());
                itm.SubItems.Add(dr[10].ToString());
                itm.SubItems.Add(dr[12].ToString());
                itm.SubItems.Add(dr[13].ToString());
                itm.SubItems.Add(dr[14].ToString());
                itm.SubItems.Add(dr[15].ToString());
                itm.SubItems.Add(dr[16].ToString());
                itm.SubItems.Add(dr[17].ToString());
                itm.SubItems.Add(dr[18].ToString());
                itm.SubItems.Add(dr[19].ToString());
                listView1.Items.Add(itm);
                allitems[i] = itm;
                
                i++;
            }
            if (comboBox1.Text == "" || comboBox1.Text == null) comboBox1.Text = "ALL";
            comboBox1_SelectedIndexChanged(null, null);
        }
        List<ListViewItem> datedlist(List<ListViewItem> itmlist)
        {
            if (checkBox1.Checked)
            {
                List<ListViewItem> returnedlist = new List<ListViewItem>();
                foreach (ListViewItem itm in itmlist)
                {
                    if (DateTime.Parse(itm.SubItems[3].Text) >= dateTimePicker1.Value && DateTime.Parse(itm.SubItems[3].Text) <= dateTimePicker2.Value)
                        returnedlist.Add(itm);
                }
                return returnedlist;
            }
            else return itmlist;
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {


            updatelist();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            selectedfilter = new List<ListViewItem>();
           
            switch (comboBox1.Text)
            {
                case "ALL":
                    
                    
                    foreach (ListViewItem itm in allitems)
                    {
                        selectedfilter.Add(itm);
                        
                    }
                    break;

                case "ACTIVE/PENDING":
                    
                    
                    foreach (ListViewItem itm in allitems)
                    {
                        if (double.Parse(itm.SubItems[17].Text) > 0 || (DateTime.Today < DateTime.Parse(itm.SubItems[4].Text))) selectedfilter.Add(itm);
                    }
                    break;

                case "NOT FILLED":
                    foreach(ListViewItem itm in allitems)
                    {
                        if (DateTime.Today < DateTime.Parse(itm.SubItems[4].Text)) selectedfilter.Add(itm);
                    }
                    break;

                case "OUTSTANDING":
                    foreach (ListViewItem itm in allitems)
                    {
                        if (double.Parse(itm.SubItems[17].Text) > 0) selectedfilter.Add(itm);
                    }
                    break;
                
            }
            
            updatelist();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                Form11 f = new Form11(false, listView1.SelectedItems[0].Text);
                f.ShowDialog();
                Form16_Load(null, null);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            updatelist();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form11 f = new Form11(true, "0");
            f.ShowDialog();
            Form16_Load(null, null);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                DialogResult dialogresult = MessageBox.Show("The Invoice and all items and payments\r\nrelated to it will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
                if (dialogresult == DialogResult.Yes)
                {
                    if (connection.State != ConnectionState.Open) connection.Open();
                    OleDbCommand command = new OleDbCommand("delete * from InvoiceHeader where InvoiceID=" + listView1.SelectedItems[0].Text, connection);
                    command.ExecuteNonQuery();
                    command.CommandText = "delete * from InvoiceDetails where InvoiceID=" + listView1.SelectedItems[0].Text;
                    command.ExecuteNonQuery();
                    command.CommandText = "delete * from Payments where InvoiceID=" + listView1.SelectedItems[0].Text;
                    command.ExecuteNonQuery();
                    string log = DateTime.Now.ToString("g") + " Invoice Deleted: ID: " + listView1.SelectedItems[0].Text + ", Customer:" + listView1.SelectedItems[0].SubItems[1].Text + ", Total:" + listView1.SelectedItems[0].SubItems[15].Text + ", payments:" + listView1.SelectedItems[0].SubItems[16].Text;
                    using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                    {
                        wr.WriteLine(log);
                    }
                    Form16_Load(null, null);
                }
            }
            else MessageBox.Show("Select an Invoice");
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            updatelist();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            updatelist();
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            button2_Click(null, null);
        }
    }
}
