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
    public partial class Form7 : Form
    {
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        int currentid;
        public Form7(string id)
        {
            currentid = int.Parse(id);
            InitializeComponent();
            
        }

        private void Form7_Load(object sender, EventArgs e)
        {
            textBox1.Text = currentid.ToString();
            connection.Open();
            OleDbCommand command = new OleDbCommand("select CustomerID, VoucherDate, CreditItem, Reason, EmployeeIDReq, ManagerApprv, DateApprv, EmployeeIDApply, DateApplied from VoucherLog where VoucherID=" + currentid, connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int customerid = reader.GetInt32(0);
            
            if (!reader.IsDBNull(1)) dateTimePicker1.Text = reader.GetString(1); else dateTimePicker1.Text = "";
            if (!reader.IsDBNull(2)) textBox5.Text = reader.GetString(2); else textBox5.Text = "";
            if (!reader.IsDBNull(3)) textBox3.Text = reader.GetString(3); else textBox3.Text = "";
            if (!reader.IsDBNull(4)) textBox6.Text = reader.GetString(4); else textBox6.Text = "";
            if (!reader.IsDBNull(5)) textBox7.Text = reader.GetString(5); else textBox7.Text = "";
            if (!reader.IsDBNull(6)) dateTimePicker2.Text = reader.GetString(6); else dateTimePicker2.Text = "";
            if (!reader.IsDBNull(7)) textBox4.Text = reader.GetString(7); else textBox4.Text = "";
            if (!reader.IsDBNull(8)) dateTimePicker3.Text = reader.GetString(8); else dateTimePicker3.Text = "";
            reader.Close();
            command.CommandText = "select FirstName, LastName from Customers where CustomerID=" + customerid;
            reader = command.ExecuteReader();
            reader.Read();
            textBox2.Text = reader.GetString(0) + " " + reader.GetString(1);
            reader.Close();
            command.Dispose();
            connection.Close();

              
        }

        private void button2_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 VoucherID FROM VoucherLog ORDER BY VoucherID DESC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int lastid = reader.GetInt32(0);
            reader.Close();
            if (currentid < lastid)
            {
                int i = 1;
                command.CommandText = "select VoucherID from VoucherLog where VoucherID=" + (currentid + i);
                reader = command.ExecuteReader();

                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select VoucherID from VoucherLog where VoucherID=" + (currentid + i);
                    reader = command.ExecuteReader();

                }

                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid + i;
                Form7_Load(null, null);


            }
            else
            {
                currentid = 0;
                int i = 1;
                command.CommandText = "select VoucherID from VoucherLog where VoucherID=" + (currentid + i);
                reader = command.ExecuteReader();

                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select VoucherID from VoucherLog where VoucherID=" + (currentid + i);
                    reader = command.ExecuteReader();

                }

                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid + i;
                Form7_Load(null, null);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 VoucherID FROM VoucherLog ORDER BY VoucherID ASC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int firstid = reader.GetInt32(0);
            reader.Close();
            if (currentid > firstid)
            {
                
                int i = 1;
                command.CommandText = "select VoucherID from VoucherLog where VoucherID=" + (currentid - i);
                reader = command.ExecuteReader();
                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select VoucherID from VoucherLog where VoucherID=" + (currentid - i);
                    reader = command.ExecuteReader();

                }
                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid - i;
                Form7_Load(null, null);

            }
            reader.Close();
            command.Dispose();
            connection.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Selected Voucher will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select VoucherID, CustomerID, CreditItem from VoucherLog where VoucherID=" + currentid, connection);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                string log = DateTime.Now.ToString("g") + ": Voucher deletet: id:" + reader.GetInt32(0) + " Customer ID: " + reader.GetInt32(1) + " Credit Item: " + reader.GetString(2) + ";";
                using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                {
                    wr.WriteLine(log);
                }

                //wr.WriteLine(log);
                reader.Close();
                command.CommandText = "delete * from VoucherLog where VoucherID=" + currentid;
                command.ExecuteNonQuery();
                command.Dispose();
                connection.Close();
                
                button2_Click(null, null);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("update VoucherLog set VoucherDate='" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "', CreditItem='" + textBox5.Text+"', Reason='" + textBox3.Text + "', EmployeeIDReq='" + textBox6.Text + "', ManagerApprv='" + textBox7.Text + "', DateApprv='" + dateTimePicker2.Value.ToString("MM/dd/yyyy") + "', EmployeeIDApply='" + textBox4.Text + "', DateApplied='" + dateTimePicker3.Value.ToString("MM/dd/yyyy") + "'", connection);
            command.ExecuteNonQuery();
            MessageBox.Show("Changes saved");
            command.Dispose();
            connection.Close();
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
