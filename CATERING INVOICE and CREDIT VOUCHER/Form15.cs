using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CATERING_INVOICE_and_CREDIT_VOUCHER
{
    public partial class Form15 : Form
    {
        public Form15(ListViewItem itm)
        {
            changeditem = itm;
            InitializeComponent();
        }
        public bool valid { get; set; }
        public ListViewItem changeditem { get; set; }

        private void Form15_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("cash");
            comboBox1.Items.Add("credit card");
            dateTimePicker1.Text = changeditem.SubItems[1].Text;
            textBox1.Text = changeditem.SubItems[2].Text;
            comboBox1.Text = changeditem.SubItems[3].Text;
            textBox2.Text = changeditem.SubItems[4].Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double a;
            if (double.TryParse(textBox1.Text, out a))
            {
                valid = true;
                changeditem.SubItems[1].Text = dateTimePicker1.Value.ToString("MM/dd/yyyy");
                changeditem.SubItems[2].Text = textBox1.Text;
                changeditem.SubItems[3].Text = comboBox1.Text;
                changeditem.SubItems[4].Text = textBox2.Text;
                this.Close();
            }
            else MessageBox.Show("insert a value");
        }
    }
}
