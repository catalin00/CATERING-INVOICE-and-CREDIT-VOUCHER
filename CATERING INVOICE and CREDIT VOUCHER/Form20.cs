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
    public partial class Form20 : Form
    {
        string path;
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        public Form20()
        {
            InitializeComponent();
        }

        private void Form20_Load(object sender, EventArgs e)
        {
            if (connection.State != ConnectionState.Open) connection.Open();
            OleDbCommand command = new OleDbCommand("select * from Settings",  connection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            if (!reader.IsDBNull(2)) photo = (byte[])reader.GetValue(2);
            if (!reader.IsDBNull(3)) textBox2.Text = reader.GetString(3); 
            if (!reader.IsDBNull(4)) textBox3.Text = reader.GetString(4);
            if (!reader.IsDBNull(5)) textBox4.Text = reader.GetString(5);
            if (!reader.IsDBNull(6)) textBox5.Text = reader.GetString(6);
            if (!reader.IsDBNull(7)) textBox6.Text = reader.GetString(7);
            if (!reader.IsDBNull(8)) textBox7.Text = reader.GetString(8);
            if (!reader.IsDBNull(9)) textBox8.Text = reader.GetString(9);
            if (!reader.IsDBNull(10)) textBox9.Text = reader.GetString(10);
            if (!reader.IsDBNull(11)) textBox10.Text = reader.GetString(11);
            if (!reader.IsDBNull(12)) textBox11.Text = reader.GetDouble(12).ToString();
            if (!reader.IsDBNull(13)) textBox12.Text = reader.GetDouble(13).ToString(); ;
            connection.Close();
        }
        byte[] photo;
        private void button2_Click(object sender, EventArgs e)
        {
            connection.Open();
            if (path != null && path!="") photo = File.ReadAllBytes(path);
            OleDbCommand cmd = new OleDbCommand("select top 1 id from settings order by id desc", connection);
            OleDbDataReader rdr = cmd.ExecuteReader();
            rdr.Read();
            int id = rdr.GetInt32(0);
            OleDbCommand command = new OleDbCommand("Update Settings set Logo= @img, CompanyName='"+textBox2.Text.Replace("'", "''")+"', Adress=@adress, City='" + textBox4.Text + "', State='" + textBox5.Text + "', ZipCode='" + textBox6.Text + "', Phone='" + textBox7.Text + "', Fax='" + textBox8.Text + "', URL='" + textBox9.Text + "', email='" + textBox10.Text + "', TaxRate=" + textBox11.Text.Replace(',', '.') + ", RackCharge=" + textBox12.Text.Replace(',', '.')+" where ID="+id, connection);
            command.Parameters.AddWithValue("@img", photo);
            command.Parameters.AddWithValue("@adress", textBox3.Text);
            command.Parameters.Add("@img",
      OleDbType.LongVarBinary, photo.Length).Value = photo;
           // command.Parameters.Add("@company", OleDbType.Char).Value = textBox2.Text;
            command.ExecuteNonQuery();
            MessageBox.Show("Updates Saved");
            this.Close();
        }

        public static byte[] GetPhoto(string filePath)
        {
            FileStream stream = new FileStream(
                filePath, FileMode.Open, FileAccess.Read);
            BinaryReader reader = new BinaryReader(stream);

            byte[] photo = reader.ReadBytes((int)stream.Length);

            reader.Close();
            stream.Close();

            return photo;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();

            DialogResult rez = openfile.ShowDialog();
            if (rez==DialogResult.OK)
            {
                path = openfile.FileName;
                textBox1.Text = path;
            }
        }
    }
}
