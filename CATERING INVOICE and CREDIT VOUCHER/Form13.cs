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
    public partial class Form13 : Form
    {
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        ListViewItem itm;
        int invID;
        public Form13(ListViewItem i, int id)
        {
            invID = id;
            itm = i;
            InitializeComponent();
        }
        public ListViewItem newitem { get; set; }
        public bool valid { get; set; }
        private void Form13_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("yes");
            comboBox1.Items.Add("no");
            textBox1.Text = itm.Text;
            textBox2.Text = itm.SubItems[1].Text;
            textBox4.Text = itm.SubItems[3].Text;
            textBox3.Text = itm.SubItems[4].Text;
            comboBox1.Text = itm.SubItems[6].Text;
            textBox5.Text = itm.SubItems[2].Text;
            connection.Open();
            if (invID != 0)
            {
                OleDbCommand command = new OleDbCommand("select DishDescription from InvoiceDetails where productID=" + textBox1.Text + " and InvoiceID=" + invID, connection);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                if (!reader.IsDBNull(0)) textBox5.Text = reader.GetString(0);
                reader.Close();
                command.Dispose();
            }
            connection.Close();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            valid = true;
            newitem = new ListViewItem(textBox1.Text);
            newitem.SubItems.Add(textBox2.Text);
            newitem.SubItems.Add(textBox5.Text);
            newitem.SubItems.Add(textBox4.Text);
            newitem.SubItems.Add(textBox3.Text);
            newitem.SubItems.Add((double.Parse(textBox3.Text) * double.Parse(textBox4.Text)).ToString());
            newitem.SubItems.Add(comboBox1.Text);
            if (invID != 0)
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("update InvoiceDetails set DishDescription='" + textBox5.Text + "' where InvoiceID=" + invID + " and productID=" + textBox1.Text, connection);
                command.ExecuteNonQuery();
                connection.Close();
            }
            this.Close();
        }
    }
}
