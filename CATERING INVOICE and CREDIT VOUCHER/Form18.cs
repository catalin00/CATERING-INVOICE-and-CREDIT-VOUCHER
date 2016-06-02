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
    public partial class Form18 : Form
    {
        bool ok = false;
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        public Form18()
        {
            InitializeComponent();
        }

        private void Form18_Load(object sender, EventArgs e)
        {
            if (connection.State != ConnectionState.Open) connection.Open();
            OleDbCommand command = new OleDbCommand("Select CustomerId from InvoiceHeader where BalanceDue>0", connection);
            OleDbDataReader reader = command.ExecuteReader();
            List<int> custlist = new List<int>();
            while (reader.Read())
            {
                OleDbCommand cmd = new OleDbCommand("select FirstName, LastName from Customers where CustomerID=" + reader.GetInt32(0), connection);
                OleDbDataReader rdr = cmd.ExecuteReader();
                rdr.Read();
                ComboBoxItem itm = new ComboBoxItem();
                itm.Text = rdr.GetString(0) + " " + rdr.GetString(1);
                itm.Value = reader.GetInt32(0);
                if (!custlist.Contains(reader.GetInt32(0)))
                {
                    custlist.Add(reader.GetInt32(0));
                    comboBox1.Items.Add(itm);
                }

            }
            connection.Close();
            ok = true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ok)
            {
                ComboBoxItem itm = (ComboBoxItem)comboBox1.SelectedItem;
                textBox1.Text = itm.Value.ToString();
                if (connection.State != ConnectionState.Open) connection.Open();
                comboBox2.Items.Clear();
                comboBox2.Text = "";
                OleDbDataAdapter adapter = new OleDbDataAdapter("select InvoiceID from InvoiceHeader where CustomerID=" + itm.Value.ToString() +" and BalanceDue>0", connection);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];
                   
                    
                    comboBox2.Items.Add(dr[0].ToString());

                }
                adapter.Dispose();
                dt.Dispose();

                connection.Close();

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" && comboBox2.Text != null)
            {
                Form11 f = new Form11(false, comboBox2.Text);
                f.Show();
                this.Close();
            }
            else MessageBox.Show("select an invoice");
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" && comboBox2.Text != null)
            {
                if (connection.State != ConnectionState.Open) connection.Open();
                OleDbCommand command = new OleDbCommand("insert into Payments (CustomerID, InvoiceID, PaymentDate, PaymentAmount, PaymentType, MicrosRefNo) values (" + textBox1.Text + ", " + comboBox2.Text + ", '" + DateTime.Today.ToString("MM/dd/yyyy") + "', " + textBox2.Text.Replace(',', '.') + ", '" + comboBox3.Text + "', '" + textBox3.Text + "')", connection);
                command.ExecuteNonQuery();
                command.CommandText = "update invoiceHeader set Payments=(Payments+" + textBox2.Text.Replace(',', '.') + "), BalanceDue=(BalanceDue-" + textBox2.Text.Replace(',', '.') + ") where InvoiceID=" + comboBox2.Text;
                command.ExecuteNonQuery();
                MessageBox.Show("Payment Saved!");
                this.Close();
            }
            else MessageBox.Show("select an invoice");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
