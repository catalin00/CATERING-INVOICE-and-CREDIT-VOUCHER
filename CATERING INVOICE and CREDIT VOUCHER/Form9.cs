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
    public partial class Form9 : Form
    {
        string taxablefood = "no";
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        public Form9()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            decimal a;
            if (!decimal.TryParse(textBox3.Text, out a)) MessageBox.Show("Insert a valid price");
            else
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("insert into FoodProducts (DishName, DishDescription, TaxableFood, UnitPrice) values ('" + textBox1.Text + "', '" + textBox2.Text + "', '" + taxablefood + "', '" + textBox3.Text + "')", connection);
                command.ExecuteNonQuery();
                command.Dispose();
                connection.Close();
                MessageBox.Show("Item Saved!");
                this.Close();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) taxablefood = "yes";
            else taxablefood = "no";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
