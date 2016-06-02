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
    public partial class Form14 : Form
    {
        public double amount { get; set; }
        public string micros { get; set; }
        public bool valid { get; set; }
        public string type { get; set; }
        
        public Form14()
        {
            valid = false;
            InitializeComponent();
        }

        private void Form14_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("cash");
            comboBox1.Items.Add("credit card");
            comboBox1.Text = "cash";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            valid = true;
            amount = double.Parse(textBox2.Text);
            micros = textBox3.Text;
            type = comboBox1.Text;
            this.Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
