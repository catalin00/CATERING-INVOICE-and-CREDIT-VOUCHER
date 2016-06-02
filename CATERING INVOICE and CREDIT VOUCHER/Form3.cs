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
    public partial class Form3 : Form
    {
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        
        public Form3()
        {
            
            InitializeComponent();
        }
        public string name { get; set; }
        public int id { get; set; }
        private void button1_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("insert into Customers (Salutation, FirstName, LastName, Title, MobileNumber, Department, CompanyName, PhoneNumber, PhoneExt, FaxNumber, EmailAdress, WebsiteURL, Adress, City, State, postalCode, caAdress, caCity, caState, caZip, Notes) values ('"+textBox5.Text+"', '"+textBox2.Text+"', '"+textBox3.Text+"', '"+textBox4.Text+"', '"+textBox6.Text+"', '"+textBox7.Text+"', '"+textBox8.Text+"', '"+textBox9.Text+"', '"+textBox10.Text+"', '"+textBox11.Text+"', '"+textBox12.Text+"', '"+textBox13.Text+"', '"+textBox14.Text+"', '"+textBox15.Text+"', '"+textBox16.Text+"', '"+textBox17.Text+"', '"+textBox18.Text+"', '"+textBox19.Text+"', '"+textBox20.Text+"', '"+textBox21.Text+"', '"+textBox22.Text+"')", connection);
            command.ExecuteNonQuery();
            name = textBox2.Text + " " + textBox3.Text;
            command.CommandText = "select top 1 CustomerID from customers order by CustomerID desc";
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            id = reader.GetInt32(0);
            command.Dispose();
            connection.Close();
            MessageBox.Show("Customer Saved");
            valid = true;
            this.Close();
        }
        public bool valid { get; set; }
        private void Form3_Load(object sender, EventArgs e)
        {
            valid = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
