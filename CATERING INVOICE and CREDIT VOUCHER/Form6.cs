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
    public partial class Form6 : Form
    {
        int customerid;
        bool ok = false;
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        public Form6()
        {
            InitializeComponent();
        }

        private void Form6_Load(object sender, EventArgs e)
        {
            connection.Open();
            
            OleDbDataAdapter adapter = new OleDbDataAdapter("select customerID, FirstName, LastName from Customers", connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            for (int i=0; i<dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ComboBoxItem itm = new ComboBoxItem();
                itm.Text = dr[1].ToString() + " " + dr[2].ToString();
                itm.Value = dr[0].ToString();
                comboBox1.Items.Add(itm);

            }
            connection.Close();
            ok = true;


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ok)
            {
                customerid = 0;
                ComboBoxItem itm = (ComboBoxItem)comboBox1.SelectedItem;
                textBox2.Text = (string)itm.Value;
                customerid = int.Parse((string)itm.Value);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (customerid != 0)
            {


                connection.Open();
                OleDbCommand command = new OleDbCommand("insert into VoucherLog (CustomerID, VoucherDate, CreditItem, Reason, EmployeeIDReq, ManagerApprv, DateApprv, EmployeeIDApply, DateApplied) values (" + customerid + ", '" + DateTime.Today.ToString("MM/dd/yyyy") + "', '" + textBox5.Text + "', '" + textBox4.Text + "', '" + textBox6.Text + "', '" + textBox7.Text + "', '" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "', '" + textBox3.Text + "', '" + dateTimePicker2.Value.ToString("MM/dd/yyyy") + "')", connection);
                command.ExecuteNonQuery();
                command.Dispose();
                connection.Close();
                MessageBox.Show("voucher credit added");
                this.Close();
            }
            else MessageBox.Show("Select a valid customer!");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var f = new Form3())
            {
                f.ShowDialog();
                if (f.valid)
                {
                    customerid = f.id;
                    textBox2.Text = customerid.ToString();
                    comboBox1.Text = f.name;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (var f = new Form2(true))
            {
                f.ShowDialog();
                if (f.valid)
                {
                    customerid = f.custid;
                    textBox2.Text = customerid.ToString();
                    comboBox1.Text = f.name;
                }
            }
        }
    }
}
