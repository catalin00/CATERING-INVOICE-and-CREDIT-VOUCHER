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
    public partial class Form4 : Form
    {
        OleDbConnection connection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");

        int currentid;
        public Form4( string id)
        {
            InitializeComponent();
            currentid = int.Parse(id);

        }

        private void Form4_Load(object sender, EventArgs e)
        {
            textBox1.Text = currentid.ToString();
            connection.Open();
            OleDbCommand command = new OleDbCommand("select * from Customers where CustomerID=" + currentid, connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            if (!reader.IsDBNull(2)) label2.Text = reader.GetString(2); else label2.Text = "";
            if (!reader.IsDBNull(3)) label2.Text = label2.Text+" "+ reader.GetString(3);
            if (!reader.IsDBNull(1)) textBox6.Text = reader.GetString(1); else textBox6.Text = "";
            if (!reader.IsDBNull(2)) textBox8.Text = reader.GetString(2); else textBox8.Text = "";
            if (!reader.IsDBNull(3)) textBox7.Text = reader.GetString(3); else textBox7.Text = "";
            if (!reader.IsDBNull(4)) textBox5.Text = reader.GetString(4); else textBox5.Text = "";
            if (!reader.IsDBNull(5)) textBox4.Text = reader.GetString(5); else textBox4.Text = "";
            if (!reader.IsDBNull(6)) textBox3.Text = reader.GetString(6); else textBox3.Text = "";
            if (!reader.IsDBNull(7)) textBox2.Text = reader.GetString(7); else textBox2.Text = "";
            if (!reader.IsDBNull(8)) textBox9.Text = reader.GetString(8); else textBox9.Text = "";
            if (!reader.IsDBNull(9)) textBox10.Text = reader.GetString(9); else textBox10.Text = "";
            if (!reader.IsDBNull(10)) textBox11.Text = reader.GetString(10); else textBox11.Text = "";
            if (!reader.IsDBNull(11)) textBox12.Text = reader.GetString(11); else textBox12.Text = "";
            if (!reader.IsDBNull(12)) textBox13.Text = reader.GetString(12); else textBox13.Text = "";
            if (!reader.IsDBNull(13)) textBox14.Text = reader.GetString(13); else textBox14.Text = "";
            if (!reader.IsDBNull(14)) textBox15.Text = reader.GetString(14); else textBox15.Text = "";
            if (!reader.IsDBNull(15)) textBox17.Text = reader.GetString(15); else textBox17.Text = "";
            if (!reader.IsDBNull(16)) textBox18.Text = reader.GetString(16); else textBox18.Text = "";
            if (!reader.IsDBNull(17)) textBox19.Text = reader.GetString(17); else textBox19.Text = "";
            if (!reader.IsDBNull(18)) textBox20.Text = reader.GetString(18); else textBox20.Text = "";
            if (!reader.IsDBNull(19)) textBox21.Text = reader.GetString(19); else textBox21.Text = "";
            if (!reader.IsDBNull(20)) textBox22.Text = reader.GetString(20); else textBox22.Text = "";
            if (!reader.IsDBNull(21)) textBox16.Text = reader.GetString(21); else textBox16.Text = "";
            reader.Close();
            command.Dispose();
            OleDbDataAdapter ada = new OleDbDataAdapter("select InvoiceID, PaymentID, PaymentDate, PaymentAmount, PaymentType from Payments where CustomerID=" + currentid.ToString(), connection);
            DataTable dt = new DataTable();
            ada.Fill(dt);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[0].ToString());
                for (int x = 1; x < 5; x++)
                {
                    listitem.SubItems.Add(dr[x].ToString());
                }

                listView1.Items.Add(listitem);
            }

            dt.Dispose();
            ada.Dispose();
            ada = new OleDbDataAdapter("select InvoiceID, InvoiceDate, CaterDate, CaterTime, OrderTypePD, Total, Payments, Notes from InvoiceHeader where CustomerID=" + currentid.ToString(), connection);
            dt = new DataTable();
            ada.Fill(dt);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[0].ToString());
                for (int x = 1; x < 8; x++)
                {
                    listitem.SubItems.Add(dr[x].ToString());
                }

                listView3.Items.Add(listitem);
            }
            dt.Dispose();
            ada.Dispose();
            ada = new OleDbDataAdapter("select VoucherID, VoucherDate, CreditItem, Reason from VoucherLog where CustomerID=" + currentid.ToString(), connection);
            dt = new DataTable();
            ada.Fill(dt);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[0].ToString());
                for (int x = 1; x < 4; x++)
                {
                    listitem.SubItems.Add(dr[x].ToString());
                }

                listView2.Items.Add(listitem);
            }
            dt.Dispose();
            ada.Dispose();
            connection.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            listView2.Items.Clear();
            listView3.Items.Clear();
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 CustomerID FROM Customers ORDER BY CustomerID DESC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int lastid = reader.GetInt32(0);
            reader.Close();
            if (currentid<lastid)
            {
                int i = 1;
                command.CommandText = "select CustomerID from Customers where CustomerID=" + (currentid + i);
                reader = command.ExecuteReader();
                
                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText= "select CustomerID from Customers where CustomerID=" + (currentid + i);
                    reader = command.ExecuteReader();
                    
                }
                
                    reader.Close();
                    command.Dispose();
                    connection.Close();
                    currentid = currentid + i;
                    Form4_Load(null, null);
                
                
            }
            else
            {
                currentid = 0;
                int i = 1;
                command.CommandText = "select CustomerID from Customers where CustomerID=" + (currentid + i);
                reader = command.ExecuteReader();

                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select CustomerID from Customers where CustomerID=" + (currentid + i);
                    reader = command.ExecuteReader();

                }

                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid + i;
                Form4_Load(null, null);

            }
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 CustomerID FROM Customers ORDER BY CustomerID ASC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int firstid = reader.GetInt32(0);
            reader.Close();
            if (currentid>firstid)
            {
                listView1.Items.Clear();
                listView2.Items.Clear();
                listView3.Items.Clear();
                int i = 1;
                command.CommandText = "select CustomerID from Customers where CustomerID=" + (currentid - i);
                reader = command.ExecuteReader();
                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select CustomerID from Customers where CustomerID=" + (currentid - i);
                    reader = command.ExecuteReader();
                    
                }
                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid - i;
                Form4_Load(null, null);

            }
            reader.Close();
            command.Dispose();
            connection.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button5_Click(null, null);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button6_Click(null, null);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("update Customers set Salutation='" + textBox6.Text + "', FirstName='" + textBox8.Text + "', LastName='" + textBox7.Text + "', Title='" + textBox5.Text + "', MobileNumber='" + textBox4.Text + "', Department='" + textBox3.Text + "', CompanyName='" + textBox2.Text + "', PhoneNumber='" + textBox9.Text + "', PhoneExt='" + textBox10.Text + "', FaxNumber='" + textBox11.Text + "', EmailAdress='" + textBox12.Text + "', WebsiteURL='" + textBox13.Text + "', Adress='" + textBox14.Text + "', City='" + textBox15.Text + "', State='" + textBox17.Text + "', PostalCode='" + textBox18.Text + "', caAdress='" + textBox19.Text + "', caCity='" + textBox20.Text + "', caState='" + textBox21.Text + "', caZip='" + textBox22.Text + "' where CustomerID=" + currentid.ToString(), connection);
            command.ExecuteNonQuery();
            MessageBox.Show("Changes saved");
            command.Dispose();
            connection.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Selected customer will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select FirstName, LastName from Customers where CustomerID=" + currentid, connection);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                string log = DateTime.Now.ToString("g") + ": customer deleted id:" + currentid + " First Name: " + reader.GetString(0) + " Last Name: " + reader.GetString(1) + ";";
                using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                {
                    wr.WriteLine(log);
                }

                //wr.WriteLine(log);
                reader.Close();
                command.CommandText = "delete * from Customers where CustomerID=" + currentid;
                command.ExecuteNonQuery();
                command.CommandText = "delete * from Voucherlog where CustomerID=" + currentid;
                command.ExecuteNonQuery();
                command.CommandText = "select InvoiceId from InvoiceHeader where CustomerID=" + currentid;
                OleDbDataReader rdr = command.ExecuteReader();
                List<int> idlist = new List<int>();
                while (rdr.Read())
                {
                    idlist.Add(rdr.GetInt32(0));
                }
                rdr.Close();
                foreach (int i in idlist)
                {
                    command.CommandText = "delete * from InvoiceDetails where invoiceID=" + i;
                    command.ExecuteNonQuery();
                }
                command.CommandText = "delete * from InvoiceHeader where CustomerID=" + currentid;
                command.ExecuteNonQuery();
                command.CommandText = "delete * from Payments where CustomerID=" + currentid;
                command.ExecuteNonQuery();
                command.Dispose();
                connection.Close();
                listView1.Items.Clear();
                button5_Click(null, null);
            }
        }
    }
}
