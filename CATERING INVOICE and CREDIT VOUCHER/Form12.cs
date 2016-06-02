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

namespace CATERING_INVOICE_and_CREDIT_VOUCHER
{
    public partial class Form12 : Form
    {
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        ListViewItem[] allitems;
        
        
        public Form12()
        {
            InitializeComponent();
            
        }
        public ListViewItem selected { get; set; }
        public double quantity { get; set; }
        public bool valid { get; set; }
        private void Form12_Load(object sender, EventArgs e)
        {
            selected = null;
            quantity = 0;
            connection.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from FoodProducts", connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            allitems = new ListViewItem[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[0].ToString());
                for (int x = 1; x < 5; x++)
                {
                    listitem.SubItems.Add(dr[x].ToString());
                }
                allitems[i] = listitem;
                listView1.Items.Add(listitem);
            }
            connection.Close();

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

        private void button1_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                valid = true;
                selected = listView1.SelectedItems[0];
                quantity = double.Parse(textBox2.Text);
                this.Close();
            }
            else MessageBox.Show("Select an item!");
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            button1_Click(null, null);
        }
    }
}
