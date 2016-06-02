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
    public partial class Form8 : Form
    {
        ListViewItem[] allitems;
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        public Form8()
        {
            InitializeComponent();
        }

        private void Form8_Load(object sender, EventArgs e)
        {
            connection.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from FoodProducts", connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            allitems = new ListViewItem[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                ListViewItem listitem = new ListViewItem(dr[0].ToString());
                for (int x = 1; x < 5; x++)
                {
                    listitem.SubItems.Add(dr[x].ToString());
                }
                listitem.SubItems[4].Text = "$ " + listitem.SubItems[4].Text;
                allitems[i] = listitem;
                listView1.Items.Add(listitem);
            }
            connection.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form9 f = new Form9();
            f.ShowDialog();
            listView1.Items.Clear();
            Form8_Load(null, null);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form10 f;
            if (listView1.SelectedItems.Count > 0) f = new Form10(listView1.SelectedItems[0].Text);
            else f = new Form10(listView1.Items[0].Text);
            f.ShowDialog();
            listView1.Items.Clear();
            Form8_Load(null, null);
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            button2_Click(null, null);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count>0)
            {

            
            DialogResult dialogResult = MessageBox.Show("Selected Product will be deleted.\r\nContinue?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("select DishName, TaxableFood, UnitPrice from FoodProducts where ProductID=" + listView1.SelectedItems[0].Text, connection);
                    OleDbDataReader reader = command.ExecuteReader();
                    reader.Read();
                    string log = DateTime.Now.ToString("g") + ": Product deleted id:" + listView1.SelectedItems[0].Text + " Dish Name: " + reader.GetString(0) + " Taxable Food: " + reader.GetString(1) + " Unit Price: " + reader.GetDouble(2).ToString() + ";";
                    using (StreamWriter wr = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "log.text"))
                    {
                        wr.WriteLine(log);
                    }

                    //wr.WriteLine(log);
                    reader.Close();
                    command.CommandText = "delete * from FoodProducts where ProductID=" + listView1.SelectedItems[0].Text;
                    command.ExecuteNonQuery();
                    command.Dispose();
                    connection.Close();
                    listView1.Items.Clear();
                    Form8_Load(null, null);
                }
                else MessageBox.Show("Select a Product!");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            if (textBox1.Text != null && textBox1.Text != "")
            {
                foreach (ListViewItem itm in allitems)
                {
                    if (itm.SubItems[1].Text.ToLower().Contains(textBox1.Text.ToLower()) ) listView1.Items.Add(itm);
                }
            }
            else listView1.Items.AddRange(allitems);
        }
    }
}
