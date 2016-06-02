using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CATERING_INVOICE_and_CREDIT_VOUCHER
{
    public partial class Form19 : Form
    {
        int currentID;
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");

        public Form19(string id)
        {
            currentID = int.Parse(id);
            InitializeComponent();
        }
        double origamount;
        private void Form19_Load(object sender, EventArgs e)
        {
            if (connection.State != ConnectionState.Open) connection.Open();
            textBox6.Text = currentID.ToString();
            OleDbCommand command = new OleDbCommand("select  CustomerID, InvoiceID, PaymentDate, PaymentAmount, PaymentType, MicrosRefNo from Payments where PaymentID=" + currentID, connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            OleDbCommand cmd = new OleDbCommand("select FirstName, lastName from customers where CustomerID=" + reader.GetInt32(0), connection);
            OleDbDataReader rdr = cmd.ExecuteReader();
            rdr.Read();
            textBox1.Text = rdr.GetString(0) + " " + rdr.GetString(1);
            textBox2.Text = reader.GetInt32(0).ToString();
            textBox3.Text = reader.GetInt32(1).ToString();
            dateTimePicker1.Text = reader.GetString(2);
            textBox4.Text = reader.GetDouble(3).ToString();
            origamount = reader.GetDouble(3);
            comboBox1.Text = reader.GetString(4);
            textBox5.Text = reader.GetString(5);
            connection.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (connection.State != ConnectionState.Open) connection.Open();
            OleDbCommand command = new OleDbCommand("update Payments set PaymentDate='" + dateTimePicker1.Text + "', PaymentAmount=" + textBox4.Text + ", PaymentType='" + comboBox1.Text + "', MicrosRefNo='" + textBox5.Text + "' where PaymentID=" + currentID, connection);
            command.ExecuteNonQuery();
            command.CommandText="update InvoiceHeader set Payments=Payments-"+origamount.ToString().Replace(',','.')+"+"+textBox4.Text.Replace(',','.')+",BalanceDue=BalanceDue+" + origamount.ToString().Replace(',', '.') + "-" + textBox4.Text.Replace(',', '.') + " where InvoiceID=" + textBox3.Text;
            command.ExecuteNonQuery();
            MessageBox.Show("Payment Saved");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (connection.State!=ConnectionState.Open) connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 PaymentID FROM Payments ORDER BY PaymentID DESC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int lastid = reader.GetInt32(0);
            reader.Close();
            if (currentID < lastid)
            {
                int i = 1;
                command.CommandText = "select PaymentID from Payments where PaymentID=" + (currentID + i);
                reader = command.ExecuteReader();

                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select PaymentID from Payments where PaymentID=" + (currentID + i);
                    reader = command.ExecuteReader();

                }

                reader.Close();
                command.Dispose();
                connection.Close();
                currentID = currentID + i;
                Form19_Load(null, null);


            }
            else
            {
                currentID = 0;
                int i = 1;
                command.CommandText = "select PaymentID from Payments where PaymentID=" + (currentID + i);
                reader = command.ExecuteReader();

                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select PaymentID from Payments where PaymentID=" + (currentID + i);
                    reader = command.ExecuteReader();

                }

                reader.Close();
                command.Dispose();
                connection.Close();
                currentID = currentID + i;
                Form19_Load(null, null);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (connection.State!=ConnectionState.Open) connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 PaymentID FROM Payments ORDER BY PaymentID ASC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int firstid = reader.GetInt32(0);
            reader.Close();
            if (currentID > firstid)
            {

                int i = 1;
                command.CommandText = "select PaymentID from Payments where PaymentID=" + (currentID - i);
                reader = command.ExecuteReader();
                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select PaymentID from Payments where PaymentID=" + (currentID - i);
                    reader = command.ExecuteReader();

                }
                reader.Close();
                command.Dispose();
                connection.Close();
                currentID = currentID - i;
                Form19_Load(null, null);

            }
            reader.Close();
            command.Dispose();
            connection.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Selected Payment will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (connection.State!=ConnectionState.Open) connection.Open();
                OleDbCommand command = new OleDbCommand("select PaymentID, PaymentAmount, InvoiceID, customerID from Payments where PaymentID=" + currentID, connection);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                double amount = reader.GetDouble(1);
                OleDbCommand cmd = new OleDbCommand("select FirstName, LastName from Customers where customerID=" + reader.GetInt32(3), connection);
                OleDbDataReader rdr = cmd.ExecuteReader();
                rdr.Read();
                string log = DateTime.Now.ToString("g") + ": Payment deletet: id:" + reader.GetInt32(0) + " Customer: " + rdr.GetString(0) + " " + rdr.GetString(1) + " Amount: " + reader.GetDouble(1) + " Invoice:" + reader.GetInt32(2) + ";";
                int invid = reader.GetInt32(2);
                using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                {
                    wr.WriteLine(log);
                }

                //wr.WriteLine(log);
                reader.Close();
                command.CommandText = "delete * from Payments where PaymentID=" + currentID;
                command.ExecuteNonQuery();
                command.CommandText = "update InvoiceHeader set Payments=Payments-" + amount.ToString().Replace(',', '.') + ", BalanceDue=BalanceDue+" + amount.ToString().Replace(',', '.')+" where InvoiceID="+invid;
                command.ExecuteNonQuery();
                command.Dispose();
                connection.Close();
                MessageBox.Show("Payment Deleted");
                button2_Click(null, null);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form11 f = new Form11(false, textBox3.Text);
            f.Show();
            this.Close();
        }
    }
}
