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
    public partial class Form10 : Form
    {
        int currentid;
        string taxable;
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        public Form10(string id)
        {
            currentid = int.Parse(id);
            InitializeComponent();
        }

        private void Form10_Load(object sender, EventArgs e)
        {
            textBox1.Text = currentid.ToString();
            connection.Open();
            OleDbCommand command = new OleDbCommand("select DishName, DishDescription, TaxableFood, UnitPrice from FoodProducts where ProductID=" + currentid, connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            if (!reader.IsDBNull(0)) textBox2.Text = reader.GetString(0); else textBox2.Text = "";
            if (!reader.IsDBNull(1)) textBox4.Text = reader.GetString(1); else textBox4.Text = "";
            if (!reader.IsDBNull(2)) taxable = reader.GetString(2); else taxable = "";
            if (!reader.IsDBNull(3)) textBox3.Text = reader.GetDouble(3).ToString(); else textBox3.Text = "";
            if (taxable.ToLower() == "yes") checkBox1.Checked = true; else checkBox1.Checked = false;
            command.Dispose();
            connection.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 ProductID FROM FoodProducts ORDER BY ProductID DESC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int lastid = reader.GetInt32(0);
            reader.Close();
            if (currentid < lastid)
            {
                int i = 1;
                command.CommandText = "select ProductID from FoodProducts where ProductID=" + (currentid + i);
                reader = command.ExecuteReader();

                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select ProductID from FoodProducts where ProductID=" + (currentid + i);
                    reader = command.ExecuteReader();

                }

                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid + i;
                Form10_Load(null, null);


            }
            else
            {
                currentid = 0;
                int i = 1;
                command.CommandText = "select ProductID from FoodProducts where ProductID=" + (currentid + i);
                reader = command.ExecuteReader();

                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select ProductID from FoodProducts where ProductID=" + (currentid + i);
                    reader = command.ExecuteReader();

                }

                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid + i;
                Form10_Load(null, null);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT TOP 1 ProductID FROM FoodProducts ORDER BY ProductID ASC", connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            int firstid = reader.GetInt32(0);
            reader.Close();
            if (currentid > firstid)
            {
                
                int i = 1;
                command.CommandText = "select ProductID from FoodProducts where ProductID=" + (currentid - i);
                reader = command.ExecuteReader();
                while (!reader.Read())
                {
                    i++;
                    reader.Close();
                    command.CommandText = "select ProductID from FoodProducts where ProductID=" + (currentid - i);
                    reader = command.ExecuteReader();

                }
                reader.Close();
                command.Dispose();
                connection.Close();
                currentid = currentid - i;
                Form10_Load(null, null);

            }
            reader.Close();
            command.Dispose();
            connection.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand("update FoodProducts set DishName='" + textBox2.Text + "', DishDescription='" + textBox4.Text + "', TaxableFood='" + taxable + "', UnitPrice=" + textBox3.Text+" where ProductID="+currentid.ToString(), connection);
            command.ExecuteNonQuery();
            command.Dispose();
            connection.Close();
            MessageBox.Show("Changes Saved!");
            button2_Click(null, null);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) taxable = "yes";
            else taxable = "no";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Selected Product will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select DishName, TaxableFood, UnitPrice from FoodProducts where ProductID=" + currentid, connection);
                OleDbDataReader reader = command.ExecuteReader();
                reader.Read();
                string log = DateTime.Now.ToString("g") + ": Product deleted id:" + currentid + " Dish Name: " + reader.GetString(0) + " Taxable Food: " + reader.GetString(1) + " Unit Price: "+reader.GetDouble(2).ToString()+";";
                using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                {
                    wr.WriteLine(log);
                }

                //wr.WriteLine(log);
                reader.Close();
                command.CommandText = "delete * from FoodProducts where ProductID=" + currentid;
                command.ExecuteNonQuery();
                command.Dispose();
                connection.Close();
                
                button2_Click(null, null);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
