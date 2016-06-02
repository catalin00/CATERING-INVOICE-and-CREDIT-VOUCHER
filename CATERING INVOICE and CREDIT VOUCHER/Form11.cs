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
using System.Drawing.Printing;
using CATERING_INVOICE_and_CREDIT_VOUCHER.Properties;

namespace CATERING_INVOICE_and_CREDIT_VOUCHER
{
    public partial class Form11 : Form
    {
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        double subtotal = 0;
        bool isprint = false;
        double taxableamount;
        double payments = 0;
        double balance = 0;
        double taxrate = 0;
        int pagenumber = 1;
        int currentid;
        bool isnew;
        bool ok = false;
        
        public Form11(bool getnew, string id)
        {
            isnew = getnew;
            if (!getnew) currentid = int.Parse(id);
            InitializeComponent();
        }

        void getpayments()
        {
            payments = 0;
            foreach (ListViewItem itm in listView2.Items)
            {
                payments = payments + double.Parse(itm.SubItems[2].Text);
            }
            calculatefields();
        }

        void fillcustomer(int id)
        {
            
            if (!(connection.State == ConnectionState.Open)) connection.Open();
            OleDbCommand command = new OleDbCommand("select FirstName, LastName, PhoneNumber, EmailAdress, Adress, City, State, PostalCode, caAdress, caCity, caState, caZip from Customers where CustomerID=" + id, connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            
            textBox5.Text = id.ToString();
            comboBox2.Text = reader.GetString(0) + " " + reader.GetString(1);
            if (!reader.IsDBNull(2)) textBox9.Text = reader.GetString(2);

            if (!reader.IsDBNull(3))  textBox11.Text = reader.GetString(3);
            if (!reader.IsDBNull(4))  textBox6.Text = reader.GetString(4);
            if (!reader.IsDBNull(5)) textBox7.Text = reader.GetString(5);
            if (!reader.IsDBNull(6)) textBox8.Text = reader.GetString(6);
            if (!reader.IsDBNull(7))  textBox10.Text = reader.GetString(7);
            if (!reader.IsDBNull(8))  textBox15.Text = reader.GetString(8);
            if (!reader.IsDBNull(9))  textBox16.Text = reader.GetString(9);
            if (!reader.IsDBNull(10))  textBox17.Text = reader.GetString(10);
            if (!reader.IsDBNull(11))  textBox18.Text = reader.GetString(11);
            reader.Close();
            //connection.Close();
        }
        void fillitems()
        {
            listView1.Items.Clear();
            taxableamount = 0;
            
            subtotal = 0;
            if (!(connection.State == ConnectionState.Open)) connection.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from InvoiceDetails where InvoiceID=" + currentid, connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[2].ToString());
                listitem.SubItems.Add(dr[3].ToString());
                listitem.SubItems.Add(dr[4].ToString());
                listitem.SubItems.Add(dr[5].ToString());
                listitem.SubItems.Add(dr[6].ToString());

                
                listitem.SubItems.Add((double.Parse(listitem.SubItems[3].Text) * double.Parse(listitem.SubItems[4].Text)).ToString());

                subtotal = subtotal + (double.Parse(listitem.SubItems[5].Text));
                listitem.SubItems.Add(dr[7].ToString());
                if (listitem.SubItems[6].Text == "yes") taxableamount = taxableamount + double.Parse(listitem.SubItems[5].Text);

                listView1.Items.Add(listitem);
            }
            
            adapter.Dispose();
            dt.Dispose();
            connection.Close();
        }

        void fillpayments()
        {
            listView2.Items.Clear();
            payments = 0;
            if (!(connection.State == ConnectionState.Open)) connection.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter("select PaymentID, PaymentDate, PaymentAmount, PaymentType, MicrosRefNo from Payments where InvoiceID=" + currentid, connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[0].ToString());
                for (int x = 1; x < 5; x++)
                {
                    listitem.SubItems.Add(dr[x].ToString());
                }
                payments = payments + double.Parse(listitem.SubItems[2].Text);
                listView2.Items.Add(listitem);
            }
            adapter.Dispose();
            dt.Dispose();
            connection.Close();
        }
        void calculatefields()
        {

            double precent = 0;
            double discount = 0;
            double tax = 0;
            double rackdep = 0;
            double total = 0;

            if (connection.State!=ConnectionState.Open) connection.Open();
           
            OleDbCommand command = new OleDbCommand("select taxrate from Settings", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            
            double.TryParse(textBox12.Text.Replace(',', '.'), out precent);
            discount = (precent / 100) * subtotal;
            tax = (taxableamount-((precent/100)*taxableamount)) * (reader.GetDouble(0)/100);
            taxrate = reader.GetDouble(0);
            reader.Close();
            command.CommandText = "select RackCharge from Settings";
            reader = command.ExecuteReader();
            reader.Read();
            if (!checkBox1.Checked)
                rackdep = int.Parse(textBox2.Text) * reader.GetDouble(0);
            else rackdep = 0;
            reader.Close();
            total = subtotal - discount + tax + rackdep;
            balance = total - payments;

            label21.Text = "$" + Math.Round(subtotal,2).ToString("0.00").Replace(',', '.');
            label22.Text = "$" + Math.Round(discount,2).ToString("0.00").Replace(',', '.');
            label23.Text = "$" + Math.Round(tax,2).ToString("0.00").Replace(',', '.');
            label24.Text = "$" + Math.Round(rackdep,2).ToString("0.00").Replace(',', '.');
            label33.Text = "$" + Math.Round(total,2).ToString("0.00").Replace(',', '.');
            label34.Text ="$"+ Math.Round(payments,2).ToString("0.00").Replace(',', '.');
            label35.Text = "$" + Math.Round(balance,2).ToString("0.00").Replace(',', '.');
            if (DateTime.Today < dateTimePicker2.Value) label32.Text = "Not Filled";
            else if (double.Parse(label35.Text.Substring(2)) > 0) label32.Text = "Outstanding";
            else label32.Text = "";
            connection.Close();

        }
        
        void getsubtotal()
        {
            subtotal = 0;
            taxableamount = 0;
            foreach (ListViewItem itm in listView1.Items)
            {
                subtotal = subtotal + (double.Parse(itm.SubItems[3].Text) * double.Parse(itm.SubItems[4].Text));
                if (itm.SubItems[6].Text == "yes") taxableamount = taxableamount + double.Parse(itm.SubItems[5].Text);
            }
            calculatefields();
        }


        private void Form11_Load(object sender, EventArgs e)
        {
            
            if (connection.State != ConnectionState.Open) connection.Open();
            comboBox2.Items.Clear();
            OleDbDataAdapter adapter = new OleDbDataAdapter("select customerID, FirstName, LastName from Customers", connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ComboBoxItem itm = new ComboBoxItem();
                itm.Text = dr[1].ToString() + " " + dr[2].ToString();
                itm.Value = dr[0].ToString();
                comboBox2.Items.Add(itm);

            }
            adapter.Dispose();
            dt.Dispose();
            
            connection.Close();

           


            if (isnew)
            {
                textBox4.Text = "new";
                button12.Enabled = false;
                button13.Enabled = false;
                button14.Enabled = false;
                label32.Text = "";

            }
            else
            {
                textBox4.Text = currentid.ToString();
                connection.Open();
                OleDbCommand command = new OleDbCommand("select CustomerID, Notes, EmployeeID, InvoiceDate, CaterDate, CaterTime, OrderTypePD, Racks, Sterno, DiscountApprv, RackReturned, RackRetDate, Discount from InvoiceHeader where InvoiceID=" + currentid, connection);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                if (!reader.IsDBNull(2)) textBox1.Text = reader.GetString(2); else textBox1.Text = "";
                dateTimePicker1.Text = reader.GetString(3);
                dateTimePicker2.Text = reader.GetString(4);
                dateTimePicker3.Text = reader.GetString(5);
                comboBox1.Text = reader.GetString(6);
                textBox2.Text = reader.GetInt32(7).ToString();
                textBox3.Text = reader.GetInt32(8).ToString();
                if (!reader.IsDBNull(9)) textBox13.Text = reader.GetString(9); else textBox13.Text = "";
                if (reader.GetString(10) == "yes") checkBox1.Checked = true; else checkBox1.Checked = false;
                if (!reader.IsDBNull(11)) dateTimePicker4.Text = reader.GetString(11);
                if (!reader.IsDBNull(12)) textBox12.Text = reader.GetDouble(12).ToString().Replace(',', '.');
                textBox14.Text = reader.GetString(1);
                fillcustomer(reader.GetInt32(0));
                //tabControl2.Enabled = false;

                button5.Enabled = false;
                button16.Enabled = false;
                comboBox2.Enabled = false;
                textBox6.Enabled = false;
                textBox7.Enabled = false;
                textBox8.Enabled = false;
                textBox9.Enabled = false;
                textBox10.Enabled = false;
                textBox11.Enabled = false;
                textBox15.Enabled = false;
                textBox16.Enabled = false;
                textBox17.Enabled = false;
                textBox18.Enabled = false;

                reader.Close();
                
                connection.Close();
                fillitems();
                fillpayments();
                calculatefields();
            }
            
            ok = true;
            connection.Close();
            connection.Close();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ok)
            {
                int customerid = 0;
                ComboBoxItem itm = (ComboBoxItem)comboBox2.SelectedItem;
                textBox5.Text = (string)itm.Value;
                customerid = int.Parse((string)itm.Value);
                fillcustomer(customerid);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            bool valid = false;
            ListViewItem itm=new ListViewItem();
            double quantity;
            using (var f = new Form12())
            {
                f.ShowDialog();
                valid = f.valid;
               

                    itm = f.selected;
                    quantity = f.quantity;
                
            }
            if (valid)
            {
                foreach (ListViewItem listitm in listView1.Items)
                {
                    if (itm.Text == listitm.Text)
                    {
                        MessageBox.Show("Select different products");
                        valid = false;
                    }
                }
            }
            if (valid)
            {
            

                ListViewItem itm2 = new ListViewItem(itm.Text);
                itm2.SubItems.Add(itm.SubItems[1].Text);
                itm2.SubItems.Add(itm.SubItems[2].Text);
                itm2.SubItems.Add(quantity.ToString().Replace(',','.'));
                itm2.SubItems.Add(itm.SubItems[4].Text.Replace(',','.'));
                itm2.SubItems.Add((quantity * double.Parse(itm.SubItems[4].Text)).ToString().Replace(',','.'));
                itm2.SubItems.Add(itm.SubItems[3].Text);
                if (!isnew)
                {
                    if (connection.State != ConnectionState.Open) connection.Open();
                    OleDbCommand command = new OleDbCommand("insert into InvoiceDetails (InvoiceID, ProductID, DishName, DishDescription, Quantity, UnitPrice, Taxable) values (" + currentid + ", "+itm2.Text+",  '"+itm2.SubItems[1].Text + "', '" + itm2.SubItems[2].Text + "'," + itm2.SubItems[3].Text + ", " + itm2.SubItems[4].Text + ", '" + itm2.SubItems[6].Text+"')", connection);
                    command.ExecuteNonQuery();
                    command.Dispose();
                    connection.Close();
                }
                
                listView1.Items.Add(itm2);
                getsubtotal();
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
               


                DialogResult dialogResult = MessageBox.Show("Selected product will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {

                    string log = DateTime.Now.ToString("g") + ": food product deleted from invoice Dish Name:" + listView1.SelectedItems[0].SubItems[1].Text + " ID: " + listView1.SelectedItems[0].Text + " quantity: " + listView1.SelectedItems[0].SubItems[3].Text + " from invoice:" + textBox4.Text + " ;";
                    using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                    {
                        wr.WriteLine(log);
                    }

                    //wr.WriteLine(log);
                    if (!isnew)
                    {
                        if (connection.State != ConnectionState.Open) connection.Open();
                        OleDbCommand command = new OleDbCommand("delete * from InvoiceDetails where InvoiceID=" + currentid + " and ProductID=" + listView1.SelectedItems[0].Text, connection);
                        command.ExecuteNonQuery();
                        command.Dispose();
                        connection.Close();
                    }
                    listView1.Items.Remove(listView1.SelectedItems[0]);
                    getsubtotal();

                }


            }
            else MessageBox.Show("select a product");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                ListViewItem itm = listView1.SelectedItems[0];
                listView1.Items.Remove(listView1.SelectedItems[0]);
                if (!isnew)
                {
                    using (var f = new Form13(itm, 0))
                    {
                        f.ShowDialog();
                        if (f.valid) itm = f.newitem;
                        if (connection.State != ConnectionState.Open) connection.Open();
                        OleDbCommand command = new OleDbCommand("update InvoiceDetails set DishName='" + itm.SubItems[1].Text + "',DishDescription='" + itm.SubItems[2].Text + "', Quantity=" + itm.SubItems[3].Text.Replace(',', '.') + ", UnitPrice=" + itm.SubItems[4].Text.Replace(',', '.') + ", Taxable='" + itm.SubItems[6].Text+"' where InvoiceID="+currentid+" and ProductID="+itm.Text, connection);
                        command.ExecuteNonQuery();
                        command.Dispose();
                        connection.Close();
                    }
                }
                else
                {
                    using (var f = new Form13(itm, currentid))
                    {
                        f.ShowDialog();
                        if (f.valid) itm = f.newitem;
                    }
                }
                listView1.Items.Add(itm);
                getsubtotal();
            }
            else MessageBox.Show("select a product");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (var f = new Form3())
            {
                f.ShowDialog();
                if (f.valid)
                {
                    textBox5.Text = f.id.ToString();
                    fillcustomer(f.id);
                }
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            calculatefields();
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            button7_Click(null, null);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ListViewItem itm = new ListViewItem("new");
            using (var f = new Form14())
            {
                bool valid = false;
                f.ShowDialog();
                valid = f.valid;
                if (valid)
                {
                    itm.SubItems.Add(DateTime.Today.ToString("MM/dd/yyyy"));
                    itm.SubItems.Add(f.amount.ToString());
                    itm.SubItems.Add(f.type);
                    itm.SubItems.Add(f.micros);
                    
                    if (!isnew)
                    {
                        if (connection.State != ConnectionState.Open) connection.Open();
                        OleDbCommand command = new OleDbCommand("insert into Payments (InvoiceID, CustomerID, PaymentDate, PaymentAmount, PaymentType, MicrosRefNo) values (" + currentid + ", " + textBox5.Text + ", '" + itm.SubItems[1].Text + "', " + itm.SubItems[2].Text.Replace(',', '.') + ", '" + itm.SubItems[3].Text + "', '" + itm.SubItems[4].Text + "')", connection);
                        command.ExecuteNonQuery();
                        command.CommandText = "update invoiceHeader set Payments=(Payments+" + itm.SubItems[2].Text.Replace(',', '.') + "), BalanceDue=(BalanceDue-" + itm.SubItems[2].Text.Replace(',', '.') + ") where InvoiceID=" + currentid;
                command.ExecuteNonQuery();
                        
                        command.CommandText = "select top 1 PaymentID from Payments order by PaymentID desc";
                        OleDbDataReader reader = command.ExecuteReader();
                        reader.Read();
                        itm.Text = reader.GetInt32(0).ToString();
                        
                    }
                    
                    listView2.Items.Add(itm);
                    getpayments();
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (listView2.SelectedItems.Count > 0)
            {
                double origvalue = double.Parse(listView2.SelectedItems[0].SubItems[2].Text);
                bool valid = false;
                using (var f = new Form15(listView2.SelectedItems[0]))
                {
                    f.ShowDialog();
                    valid = f.valid;
                    if (valid)
                    {
                       
                        listView2.Items.Remove(listView2.SelectedItems[0]);
                        listView2.Items.Add(f.changeditem);
                        if (f.changeditem.Text!="new")
                        {
                            if (connection.State != ConnectionState.Open) connection.Open();
                            OleDbCommand command = new OleDbCommand("update Payments set PaymentDate='" + f.changeditem.SubItems[1].Text + "', PaymentAmount='" + f.changeditem.SubItems[2].Text.Replace(',', '.') + "', PaymentType='" + f.changeditem.SubItems[3].Text + "', MicrosRefNo='" + f.changeditem.SubItems[4].Text + "' where PaymentID=" + f.changeditem.Text, connection);
                            command.ExecuteNonQuery();
                            OleDbCommand cmd = new OleDbCommand("update invoiceHeader set Payments=(Payments-" + origvalue.ToString().Replace(',', '.') + "+" + f.changeditem.SubItems[2].Text.Replace(',', '.') + "), BalanceDue=(BalanceDue+" + origvalue + "-" + f.changeditem.SubItems[2].Text.Replace(',', '.') + ") where InvoiceID=" + textBox4.Text, connection);
                            //command.CommandText = "update invoiceHeader set Payments=(Payments-"+origvalue.ToString().Replace(',','.')+"+" + f.changeditem.SubItems[2].Text.Replace(',', '.') + "), BalanceDue=(BalanceDue+"+origvalue+"-" + f.changeditem.SubItems[2].Text.Replace(',', '.') + ") where InvoiceID=" + currentid;
                            
                            cmd.ExecuteNonQuery();
                            command.Dispose();
                            connection.Close();
                        }

                        getpayments();
                    }
                }
            }
            else MessageBox.Show("select a payment");
        }

        private void listView2_DoubleClick(object sender, EventArgs e)
        {
            button10_Click(null, null);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (listView2.SelectedItems.Count > 0)
            {
                DialogResult dialogResult = MessageBox.Show("Selected payment will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {

                    string log = DateTime.Now.ToString("g") + ": payment deleted from invoice MicrosRefNo:" + listView2.SelectedItems[0].SubItems[4].Text + " ID: " + listView2.SelectedItems[0].Text + " amount: " + listView2.SelectedItems[0].SubItems[2].Text + " from invoice:" + textBox4.Text + " ;";
                    using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                    {
                        wr.WriteLine(log);
                    }

                    //wr.WriteLine(log);
                    if (listView2.SelectedItems[0].Text!="new")
                    {
                        if (connection.State != ConnectionState.Open) connection.Open();
                        OleDbCommand command = new OleDbCommand("delete * from Payments where PaymentID=" + listView2.SelectedItems[0].Text, connection);
                        command.CommandText = "update invoiceHeader set Payments=(Payments-" + listView2.SelectedItems[0].SubItems[2].Text + "), BalanceDue=(BalanceDue+" + listView2.SelectedItems[0].SubItems[2].Text + ") where InvoiceID=" + currentid;
                        command.ExecuteNonQuery();
                        command.CommandText = "delete * from Payments where PaymentID=" + listView2.SelectedItems[0].Text;
                        command.ExecuteNonQuery();
                        command.Dispose();
                        connection.Close();
                    }
                    listView2.Items.Remove(listView2.SelectedItems[0]);

                    getpayments();

                }
            }
            else MessageBox.Show("select a payment");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool valid = true;
            if (textBox5.Text==null || textBox5.Text=="")
            {
                MessageBox.Show("Customer not selected");
                valid = false;
            }
            if (valid)
            {


                if (isnew)
                {
                    if (connection.State != ConnectionState.Open) connection.Open();
                    string rackret;
                    if (checkBox1.Checked) rackret = "yes"; else rackret = "no";
                    OleDbCommand command = new OleDbCommand("insert into InvoiceHeader (CustomerID, EmployeeID, InvoiceDate, CaterDate, CaterTime, OrderTypePD, Racks, Sterno, Subtotal, DiscountApprv, Discount, SalesTax, RackDeposit, RackReturned, RackRetDate, Total, Payments, BalanceDue, Notes) values (" + textBox5.Text + ", '" + textBox1.Text + "', '" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "', '" + dateTimePicker2.Value.ToString("MM/dd/yyyy") + "', '" + dateTimePicker3.Value.ToString("t") + "', '" + comboBox1.Text + "', " + textBox2.Text + ", " + textBox3.Text + ", " + label21.Text.Substring(2) + ", '" + textBox13.Text + "', " + textBox12.Text + ", " + label23.Text.Substring(2) + ", " + label24.Text.Substring(2) + ", '" + rackret + "', '" + dateTimePicker4.Value.ToString("MM/dd/yyyy") + "', " + label33.Text.Substring(2) + ", " + label34.Text.Substring(2) + ", " + label35.Text.Substring(2) + ", '" + textBox14.Text + "')", connection);
                    command.ExecuteNonQuery();
                    command.CommandText = "select top 1 InvoiceID from InvoiceHeader order by InvoiceID desc";
                    OleDbDataReader reader = command.ExecuteReader();
                    reader.Read();
                    textBox4.Text = reader.GetInt32(0).ToString();
                    currentid = reader.GetInt32(0);
                    reader.Close();
                    command.Dispose();
                    foreach (ListViewItem itm in listView1.Items)
                    {
                        OleDbCommand cmd = new OleDbCommand("insert into InvoiceDetails (InvoiceId, ProductID, DishName, DishDescription, Quantity, UnitPrice, Taxable) values (" + currentid + ", " + itm.Text + ", '" + itm.SubItems[1].Text + "', '" + itm.SubItems[2].Text + "' , " + itm.SubItems[3].Text.Replace(',', '.') + ", " + itm.SubItems[4].Text.Replace(',', '.') + ", '" + itm.SubItems[6].Text + "')", connection);
                        cmd.ExecuteNonQuery();
                    }
                    foreach (ListViewItem itm in listView2.Items)
                    {
                        OleDbCommand cmd = new OleDbCommand("insert into Payments (InvoiceID, CustomerID, PaymentDate, PaymentAmount, PaymentType, MicrosRefNo) values (" + currentid + ", " + textBox5.Text + ", '" + itm.SubItems[1].Text + "', " + itm.SubItems[2].Text.Replace(',', '.') + ", '" + itm.SubItems[3].Text + "', '" + itm.SubItems[4].Text + "')", connection);
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "select top 1 PaymentID from Payments order by PaymentID desc";
                        OleDbDataReader rdr = cmd.ExecuteReader();
                        rdr.Read();
                        itm.Text = rdr.GetInt32(0).ToString();
                    }
                    tabControl2.Enabled = false;
                    isnew = false;
                    button2.Enabled = true;
                    connection.Close();
                    MessageBox.Show("Invoice Saved!");


                }
                else
                {
                    if (connection.State != ConnectionState.Open) connection.Open();
                    string rackret;
                    if (checkBox1.Checked) rackret = "yes"; else rackret = "no";
                    OleDbCommand command = new OleDbCommand("update InvoiceHeader set EmployeeId='" + textBox1.Text + "', InvoiceDate='" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "', CaterDate='" + dateTimePicker2.Value.ToString("MM/dd/yyyy") + "', CaterTime='" + dateTimePicker3.Value.ToString("t") + "', OrderTypePD='" + comboBox1.Text + "', Racks=" + textBox2.Text + ", Sterno=" + textBox3.Text + ", Subtotal=" + label21.Text.Substring(2) + ", DiscountApprv='" + textBox13.Text + "', Discount=" + textBox12.Text.Replace(',', '.') + ", SalesTax=" + label23.Text.Substring(2) + ", RackDeposit=" + label24.Text.Substring(2) + ", RackReturned='" + rackret + "', RackRetDate='" + dateTimePicker4.Value.ToString("MM/dd/yyyy") + "', Total=" + label33.Text.Substring(2) + ", Payments=" + label34.Text.Substring(2) + ", BalanceDue=" + label35.Text.Substring(2) + ", Notes='" + textBox14.Text + "' where InvoiceID=" + currentid, connection);
                    string commstr = command.CommandText;
                    string saletax = label23.Text;
                    command.ExecuteNonQuery();
                    connection.Close();
                    if (!isprint)
                        MessageBox.Show("Invoice Saved!");
                    else isprint = false;
                }
            }
        }
        Image img;
        Image paid;
        List<ListViewItem> itemlist = new List<ListViewItem>();
        
        private void button2_Click(object sender, EventArgs e)
        {
            isprint = true;
            button1_Click(null, null);
            if (connection.State != ConnectionState.Open) connection.Open();
            OleDbCommand command = new OleDbCommand("select Logo, Paidico from Settings", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            byte[] a = (byte[]) reader["Logo"];
            byte[] b = (byte[])reader["Paidico"];
            //Byte[] a = (Byte[])command.ExecuteScalar();
            connection.Close();
            foreach (ListViewItem itm in listView1.Items)
            {
                itemlist.Add(itm);
            }
            itemlist.Reverse();
            img = byteArrayToImage(a);
            paid = (Image)Resources.ResourceManager.GetObject("paid");
            paid = (Image)new Bitmap(paid, new Size(120, 80));
            
            PrintDocument document = new PrintDocument();
            document.PrintPage += Pd_PrintPage;
            PrintDialog dialog = new PrintDialog();
            dialog.Document = document;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                pagenumber = 1;
                document.Print();
            }
        }


        public Image byteArrayToImage(byte[] byteArrayIn)
        {

            System.Drawing.ImageConverter converter = new System.Drawing.ImageConverter();
            Image img = (Image)converter.ConvertFrom(byteArrayIn);

            return img;
        }


        float yline;

        float xline1;
        float xline2;
        float xline3;
        float xline4;
        float xline5;
        //List<ListViewItem> itemlist;
        private void Pd_PrintPage(object sender, PrintPageEventArgs e)
        {

            if (pagenumber != 1) yline = 10;

            e.Graphics.DrawString("Page " + pagenumber, new Font("Arial", 9), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Height - 15, 10);

            if (pagenumber == 1)
            {
                if (connection.State != ConnectionState.Open) connection.Open();
                OleDbCommand command = new OleDbCommand("select CompanyName, Adress, City, State, ZipCode, Phone, Fax, URL from Settings", connection);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                e.Graphics.DrawString(reader.GetString(0), new Font("Arial Black", 15), new SolidBrush(Color.DarkRed), 10, 10);
                e.Graphics.DrawString("Invoice", new Font("Arial Black", 15), new SolidBrush(Color.DarkRed), e.Graphics.VisibleClipBounds.Width - 100, 10);
                e.Graphics.DrawLine(new Pen(Color.DarkRed, 4), 0, 50, e.PageBounds.Width, 50);
                e.Graphics.DrawImage(img, e.Graphics.VisibleClipBounds.Width / 2 - img.Width / 2, 65);
                //e.Graphics.DrawLine(new Pen (Color.Black), 0, 65+img.Height, e.Graphics.VisibleClipBounds.Width, 65+img.Height);
                e.Graphics.DrawString(reader.GetString(1), new Font("Arial", 11), new SolidBrush(Color.Black), 10, 59);
                e.Graphics.DrawString(reader.GetString(2) + ", " + reader.GetString(3) + " " + reader.GetString(4), new Font("Arial", 11), new SolidBrush(Color.Black), 10, 76);
                e.Graphics.DrawString("Phone: " + reader.GetString(5), new Font("Arial", 11), new SolidBrush(Color.Black), 10, 93);
                e.Graphics.DrawString("Fax: " + reader.GetString(6), new Font("Arial", 11), new SolidBrush(Color.Black), 10, 110);
                e.Graphics.DrawString("Website: " + reader.GetString(7), new Font("Arial", 11), new SolidBrush(Color.Black), 10, 127);
                e.Graphics.DrawString(textBox7.Text+", "+textBox8.Text+" "+textBox10.Text, new Font("Arial", 11), new SolidBrush(Color.Black), 10, 65 + img.Height - 11);
                e.Graphics.DrawString(textBox6.Text, new Font("Arial", 11), new SolidBrush(Color.Black), 10, 65 + img.Height - 28);
                e.Graphics.DrawString(comboBox2.Text, new Font("Arial", 11), new SolidBrush(Color.Black), 10, (65 + img.Height) - 45);
                e.Graphics.DrawString("Customer Info:", new Font("Arial Black", 11), new SolidBrush(Color.Black), 10, (65 + img.Height) - 62);
                e.Graphics.DrawString("Invoice No: " + currentid, new Font("Arial", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10, 59);
                e.Graphics.DrawString("Date Order Taken: " + dateTimePicker1.Text, new Font("Arial", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10, 76);
                e.Graphics.DrawString("Employee: " + textBox1.Text, new Font("Arial", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10, 93);
                if (comboBox1.Text == "delivery")
                {
                    e.Graphics.DrawString(textBox16.Text + ", " + textBox17.Text + " " + textBox18.Text, new Font("Arial", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10, 65 + img.Height - 28);
                    e.Graphics.DrawString(textBox15.Text, new Font("Arial", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10, 65 + img.Height - 45);
                    e.Graphics.DrawString("Deliver To:", new Font("Arial Black", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10, 65 + img.Height - 62);
                }
                e.Graphics.DrawLine(new Pen(Color.DarkRed, 3), 0, 65 + img.Height + 15, e.Graphics.VisibleClipBounds.Width, 65 + img.Height + 15);
                e.Graphics.DrawString("Date", new Font("Arial Black", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 6 - 5, 65 + img.Height + 25);
                e.Graphics.DrawString("Time", new Font("Arial Black", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 3 + 13, 65 + img.Height + 25);
                e.Graphics.DrawString("Order Type", new Font("Arial Black", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + 10, 65 + img.Height + 25);
                e.Graphics.DrawString("Racks", new Font("Arial Black", 11), new SolidBrush(Color.Black), 2 * e.Graphics.VisibleClipBounds.Width / 3 + 15, 65 + img.Height + 25);
                e.Graphics.DrawString("Sterno", new Font("Arial Black", 11), new SolidBrush(Color.Black), 5 * e.Graphics.VisibleClipBounds.Width / 6 + 10, 65 + img.Height + 25);
                e.Graphics.DrawString(dateTimePicker2.Value.ToLongDateString(), new Font("Arial", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 6 - 50, 65 + img.Height + 43);
                e.Graphics.DrawString(dateTimePicker3.Text, new Font("Arial", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 3 + 6, 65 + img.Height + 43);
                e.Graphics.DrawString(comboBox1.Text, new Font("Arial", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + 15, 65 + img.Height + 43);
                e.Graphics.DrawString(textBox2.Text, new Font("Arial", 11), new SolidBrush(Color.Black), 2 * e.Graphics.VisibleClipBounds.Width / 3 + 23, 65 + img.Height + 43);
                e.Graphics.DrawString(textBox3.Text, new Font("Arial", 11), new SolidBrush(Color.Black), 5 * e.Graphics.VisibleClipBounds.Width / 6 + 18, 65 + img.Height + 43);



                e.Graphics.DrawLine(new Pen(Color.DarkRed, 3), 0, 65 + img.Height + 64, e.Graphics.VisibleClipBounds.Width, 65 + img.Height + 64);
                e.Graphics.DrawLine(new Pen(Color.Black), 1, 65 + img.Height + 67, 1, 65 + img.Height + 95);
                e.Graphics.DrawString("Qty", new Font("Arial Black", 11), new SolidBrush(Color.Black), 4, 65 + img.Height + 70);
                e.Graphics.DrawLine(new Pen(Color.Black), 41, 65 + img.Height + 67, 41, 65 + img.Height + 95);
                e.Graphics.DrawString("Description", new Font("Arial Black", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 4 + 24, 65 + img.Height + 70);
                e.Graphics.DrawLine(new Pen(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10, 65 + img.Height + 67, e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10, 65 + img.Height + 95);
                e.Graphics.DrawString("Rate", new Font("Arial Black", 11), new SolidBrush(Color.Black), e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10 + 30, 65 + img.Height + 70);
                e.Graphics.DrawLine(new Pen(Color.Black), (e.Graphics.VisibleClipBounds.Width - (e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10)) / 2 + (e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10), 65 + img.Height + 67, (e.Graphics.VisibleClipBounds.Width - (e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10)) / 2 + (e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10), 65 + img.Height + 95);
                e.Graphics.DrawString("Amount", new Font("Arial Black", 11), new SolidBrush(Color.Black), (e.Graphics.VisibleClipBounds.Width - (e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10)) / 2 + (e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10) + 20, 65 + img.Height + 70);
                e.Graphics.DrawLine(new Pen(Color.Black), e.Graphics.VisibleClipBounds.Width - 5, 65 + img.Height + 67, e.Graphics.VisibleClipBounds.Width - 5, 65 + img.Height + 95);
                e.Graphics.DrawLine(new Pen(Color.Black), 1, 65 + img.Height + 95, e.Graphics.VisibleClipBounds.Width - 5, 65 + img.Height + 95);
                yline = 65 + img.Height + 95;
                xline1 = 1;
                xline2 = 41;
                xline3 = e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10;
                xline4 = (e.Graphics.VisibleClipBounds.Width - (e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10)) / 2 + (e.Graphics.VisibleClipBounds.Width / 2 + img.Width / 2 + 10);
                xline5 = e.Graphics.VisibleClipBounds.Width - 5;
            }
            pagenumber++;
            
            foreach(ListViewItem itm in itemlist.Reverse<ListViewItem>())
            {
                
                if ((yline + 24) < e.Graphics.VisibleClipBounds.Height)
                {
                    e.Graphics.DrawLine(new Pen(Color.Black), xline1, yline, xline5, yline);
                    e.Graphics.DrawLine(new Pen(Color.Black), xline1, yline, xline1, yline + 24);
                    e.Graphics.DrawString(Math.Round(double.Parse(itm.SubItems[3].Text), 2).ToString(), new Font("Arial", 11), new SolidBrush(Color.Black), xline1 + 2, yline + 4);
                    e.Graphics.DrawLine(new Pen(Color.Black), xline2, yline, xline2, yline + 24);
                    string name = itm.SubItems[1].Text;
                    if (connection.State != ConnectionState.Open) connection.Open();
                    OleDbCommand cmd = new OleDbCommand("select DishDescription from InvoiceDetails where InvoiceID=" + currentid + " and ProductID=" + itm.Text, connection);
                    OleDbDataReader rdr = cmd.ExecuteReader();
                    rdr.Read();
                    if (!rdr.IsDBNull(0)) name = name + " (" + rdr.GetString(0) + ")";
                    if (itm.SubItems[6].Text != "yes") name = "* " + name;

                    string textramas = name;
                    string textcurent = textramas;
                    int linii = textramas.Length / 62+1;
                    for (int i=1; i<=linii; i++)
                    {
                        if (textramas.Length > 62)
                        {
                            int pos = 62;
                            while (textramas[pos] != ' ') pos--;
                            textcurent = textramas.Substring(0, pos);
                            textramas = textramas.Substring(pos);
                        }
                        else textcurent = textramas;
                        e.Graphics.DrawString(textcurent, new Font("Arial", 11), new SolidBrush(Color.Black), xline2 + 2, yline + 4);
                        e.Graphics.DrawLine(new Pen(Color.Black), xline3, yline, xline3, yline + 24);
                        e.Graphics.DrawLine(new Pen(Color.Black), xline4, yline, xline4, yline + 24);
                        e.Graphics.DrawLine(new Pen(Color.Black), xline5, yline, xline5, yline + 24);
                        e.Graphics.DrawLine(new Pen(Color.Black), xline1, yline, xline1, yline + 24);
                        e.Graphics.DrawLine(new Pen(Color.Black), xline2, yline, xline2, yline + 24);
                        if (i < linii) yline = yline + 24;
                    }




                    //e.Graphics.DrawString(name, new Font("Arial", 11), new SolidBrush(Color.Black), xline2 + 2, yline + 3);
                    e.Graphics.DrawLine(new Pen(Color.Black), xline3, yline, xline3, yline + 24);
                    e.Graphics.DrawString("$" + Math.Round(double.Parse(itm.SubItems[4].Text), 2).ToString("0.00"), new Font("Arial", 11), new SolidBrush(Color.Black), xline4-3-(e.Graphics.MeasureString("$" + Math.Round(double.Parse(itm.SubItems[4].Text), 2).ToString("0.00"), new Font("Arial", 11)).Width) , yline + 4);
                    e.Graphics.DrawLine(new Pen(Color.Black), xline4, yline, xline4, yline + 24);
                    e.Graphics.DrawString("$" + Math.Round(double.Parse(itm.SubItems[5].Text), 2).ToString("0.00"), new Font("Arial", 11), new SolidBrush(Color.Black), xline5 - 3 - (e.Graphics.MeasureString("$" + Math.Round(double.Parse(itm.SubItems[5].Text), 2).ToString("0.00"), new Font("Arial", 11)).Width), yline + 4);
                    e.Graphics.DrawLine(new Pen(Color.Black), xline5, yline, xline5, yline + 24);
                    e.Graphics.DrawLine(new Pen(Color.Black), xline1, yline + 24, xline5, yline + 24);
                    yline = yline + 24;
                    itemlist.Remove(itm);
                }
                else
                {
                    e.HasMorePages = true;
                    break;

                }
                
            }
            
            

            if (e.HasMorePages==false)
            {
                if (yline > e.Graphics.VisibleClipBounds.Height - 150) e.HasMorePages = true;
                else
                {
                    e.Graphics.DrawString("* Indicates non-taxable item", new Font("Arial", 9), new SolidBrush(Color.Black), 5, e.Graphics.VisibleClipBounds.Height - 150);
                    e.Graphics.DrawLine(new Pen(Color.Black, 2), 0, e.Graphics.VisibleClipBounds.Height - 135, e.Graphics.VisibleClipBounds.Width, e.Graphics.VisibleClipBounds.Height - 135);
                    e.Graphics.DrawString("Subtotal", new Font("Arial", 10), new SolidBrush(Color.Black), xline3-100, e.Graphics.VisibleClipBounds.Height - 115);
                    e.Graphics.DrawString("Discount", new Font("Arial", 10), new SolidBrush(Color.Black), xline3-100, e.Graphics.VisibleClipBounds.Height - 100);
                    e.Graphics.DrawString("Sales Tax" + taxrate + "%)", new Font("Arial", 10), new SolidBrush(Color.Black), xline3-100, e.Graphics.VisibleClipBounds.Height - 85);
                    e.Graphics.DrawString("Racks Deposit", new Font("Arial", 10), new SolidBrush(Color.Black), xline3 - 100, e.Graphics.VisibleClipBounds.Height - 70);
                    e.Graphics.DrawString("Total", new Font("Arial", 10), new SolidBrush(Color.Black), xline3 - 100, e.Graphics.VisibleClipBounds.Height - 55);
                    e.Graphics.DrawString("Paid", new Font("Arial", 10), new SolidBrush(Color.Black), xline3 - 100, e.Graphics.VisibleClipBounds.Height - 40);
                    e.Graphics.DrawString("Balance Due", new Font("Arial", 11), new SolidBrush(Color.Red), xline3 - 100, e.Graphics.VisibleClipBounds.Height - 25);


                    
                    e.Graphics.DrawString(label21.Text, new Font("Arial", 10), new SolidBrush(Color.Black), xline4 + 20-(e.Graphics.MeasureString(label21.Text, new Font("Arial", 10)).Width), e.Graphics.VisibleClipBounds.Height - 115);
                    e.Graphics.DrawString(textBox12.Text+"%", new Font("Arial", 10), new SolidBrush(Color.Black), xline4 + 20 - (e.Graphics.MeasureString(textBox12.Text+"%", new Font("Arial", 10)).Width), e.Graphics.VisibleClipBounds.Height - 100);
                    e.Graphics.DrawString(label23.Text, new Font("Arial", 10), new SolidBrush(Color.Black), xline4 + 20 - (e.Graphics.MeasureString(label23.Text, new Font("Arial", 10)).Width), e.Graphics.VisibleClipBounds.Height - 85);
                    e.Graphics.DrawString(label24.Text, new Font("Arial", 10), new SolidBrush(Color.Black), xline4 + 20 - (e.Graphics.MeasureString(label24.Text, new Font("Arial", 10)).Width), e.Graphics.VisibleClipBounds.Height - 70);
                    e.Graphics.DrawString(label33.Text, new Font("Arial", 10), new SolidBrush(Color.Black), xline4 + 20 - (e.Graphics.MeasureString(label33.Text, new Font("Arial", 10)).Width), e.Graphics.VisibleClipBounds.Height - 55);
                    e.Graphics.DrawString(label34.Text, new Font("Arial", 10), new SolidBrush(Color.Black), xline4 + 20 - (e.Graphics.MeasureString(label34.Text, new Font("Arial", 10)).Width), e.Graphics.VisibleClipBounds.Height - 40);
                    e.Graphics.DrawString(label35.Text, new Font("Arial", 11), new SolidBrush(Color.Red), xline4 + 20 - (e.Graphics.MeasureString(label35.Text, new Font("Arial", 11)).Width), e.Graphics.VisibleClipBounds.Height - 25);





                    if (textBox14.Text != "" && textBox14.Text != null)
                    {
                        RectangleF rectF1 = new RectangleF(10, e.Graphics.VisibleClipBounds.Height - 130, xline3 - 95, e.Graphics.VisibleClipBounds.Height - 15);
                        e.Graphics.DrawString("Notes: " + textBox14.Text, new Font("Arial", 10), new SolidBrush(Color.Black), rectF1);
                        e.Graphics.DrawRectangle(Pens.White, Rectangle.Round(rectF1));
                    }





                    if (double.Parse(label35.Text.Substring(2)) <= 0) e.Graphics.DrawImage(paid, e.Graphics.VisibleClipBounds.Width / 4 + 45, e.Graphics.VisibleClipBounds.Height - 100);
                }
            }

            if (e.HasMorePages) e.Graphics.DrawString("page " +( pagenumber-1) + " next page>>", new Font("Arial", 8), new SolidBrush(Color.Black), 5, e.Graphics.VisibleClipBounds.Height-5);
            else e.Graphics.DrawString("page " + (pagenumber-1) + " last page", new Font("Arial", 8), new SolidBrush(Color.Black), 5, e.Graphics.VisibleClipBounds.Height -15);


        }

        private void button14_Click(object sender, EventArgs e)
        {
            DialogResult dialogresult = MessageBox.Show("The Invoice and all tje items and payments\r\nrelated to it will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
            if (dialogresult==DialogResult.Yes)
            {
                if (connection.State != ConnectionState.Open) connection.Open();
                OleDbCommand command = new OleDbCommand("delete * from InvoiceHeader where InvoiceID=" + currentid, connection);
                command.ExecuteNonQuery();
                command.CommandText = "delete * from InvoiceDetails where InvoiceID=" + currentid;
                command.ExecuteNonQuery();
                command.CommandText = "delete * from Payments where InvoiceID=" + currentid;
                command.ExecuteNonQuery();
                string log = DateTime.Now.ToString("g") + " Invoice Deleted: ID: " + currentid + ", Customer:" + comboBox2.Text + ", Total:" + label33.Text + ", payments:" + label34.Text;
                using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                {
                    wr.WriteLine(log);
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (connection.State != ConnectionState.Open) connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 InvoiceID FROM InvoiceHeader ORDER BY InvoiceID DESC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int lastid = reader.GetInt32(0);
            reader.Close();
            if (currentid < lastid)
            {
                int i = 1;
                command.CommandText = "select InvoiceID from InvoiceHeader where InvoiceID=" + (currentid + i);
                reader = command.ExecuteReader();

                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select InvoiceID from InvoiceHeader where InvoiceID=" + (currentid + i);
                    reader = command.ExecuteReader();

                }

                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid + i;
                Form11_Load(null, null);


            }
            else
            {
                currentid = 0;
                int i = 1;
                command.CommandText = "select InvoiceID from InvoiceHeader where InvoiceID=" + (currentid + i);
                reader = command.ExecuteReader();

                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select InvoiceID from InvoiceHeader where InvoiceID=" + (currentid + i);
                    reader = command.ExecuteReader();

                }

                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid + i;
                Form11_Load(null, null);

            }
        }
       
        private void button3_Click(object sender, EventArgs e)
        {
            Form11_Load(null, null);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 InvoiceID FROM InvoiceHeader ORDER BY InvoiceID ASC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int firstid = reader.GetInt32(0);
            reader.Close();
            if (currentid > firstid)
            {
                
                int i = 1;
                command.CommandText = "select InvoiceID from InvoiceHeader where InvoiceID=" + (currentid - i);
                reader = command.ExecuteReader();
                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select InvoiceID from InvoiceHeader where InvoiceID=" + (currentid - i);
                    reader = command.ExecuteReader();

                }
                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid - i;
                Form11_Load(null, null);

            }
            reader.Close();
            command.Dispose();
            connection.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            using (var f = new Form2(true))
            {
                f.ShowDialog();
                if (f.valid)
                {
                    textBox5.Text = f.custid.ToString();
                    fillcustomer(f.custid);
                }
            }
        }
    }
}
