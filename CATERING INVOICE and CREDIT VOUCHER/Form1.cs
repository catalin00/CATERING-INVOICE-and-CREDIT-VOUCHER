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
    public partial class Form1 : Form
    {
        string ConnStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+AppDomain.CurrentDomain.BaseDirectory+";Jet OLEDB:Database";
        public Form1()
        {
            InitializeComponent();
        }

        private void pictureBox1_MouseEnter(object sender, EventArgs e)
        {
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.Cursor = Cursors.Hand;
        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox1.Cursor = Cursors.Default;
        }

        private void pictureBox2_MouseEnter(object sender, EventArgs e)
        {
            pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox2.Cursor = Cursors.Hand;
        }

        private void pictureBox2_MouseLeave(object sender, EventArgs e)
        {
            pictureBox2.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox2.Cursor = Cursors.Default;
        }

        private void pictureBox4_MouseEnter(object sender, EventArgs e)
        {
            pictureBox4.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox4.Cursor = Cursors.Hand;
        }

        private void pictureBox4_MouseLeave(object sender, EventArgs e)
        {
            pictureBox4.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox4.Cursor = Cursors.Default;
        }

        private void pictureBox5_MouseEnter(object sender, EventArgs e)
        {
            pictureBox5.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox5.Cursor = Cursors.Hand;
        }

        private void pictureBox5_MouseLeave(object sender, EventArgs e)
        {
            pictureBox5.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox5.Cursor = Cursors.Default;
        }

        private void pictureBox6_MouseEnter(object sender, EventArgs e)
        {
            pictureBox6.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox6.Cursor = Cursors.Hand;
        }

        private void pictureBox6_MouseLeave(object sender, EventArgs e)
        {
            pictureBox6.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox6.Cursor = Cursors.Default;
        }

        private void pictureBox9_MouseEnter(object sender, EventArgs e)
        {
            pictureBox9.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox9.Cursor = Cursors.Hand;
        }

        private void pictureBox9_MouseLeave(object sender, EventArgs e)
        {
            pictureBox9.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox9.Cursor = Cursors.Default;
        }

        private void pictureBox8_MouseEnter(object sender, EventArgs e)
        {
            pictureBox8.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox8.Cursor = Cursors.Hand;
        }

        private void pictureBox8_MouseLeave(object sender, EventArgs e)
        {
            pictureBox8.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox8.Cursor = Cursors.Default;
        }

        private void pictureBox7_MouseEnter(object sender, EventArgs e)
        {
            pictureBox7.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox7.Cursor = Cursors.Hand;
        }

        private void pictureBox7_MouseLeave(object sender, EventArgs e)
        {
            pictureBox7.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox7.Cursor = Cursors.Default;
        }

        private void pictureBox10_MouseEnter(object sender, EventArgs e)
        {
            pictureBox10.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox10.Cursor = Cursors.Hand;
        }

        private void pictureBox10_MouseLeave(object sender, EventArgs e)
        {
            pictureBox10.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox10.Cursor = Cursors.Default;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2(false);
                f.Show();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            Form5 f = new Form5();
            f.Show();
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            Form8 f = new Form8();
            f.Show();
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            Form11 f = new Form11(true, "1");
            f.Show();
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            Form16 f = new Form16();
            f.Show();
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            Form17 f = new Form17();
            f.Show();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form20 f = new Form20();
            f.ShowDialog();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Form21 f = new Form21();
            f.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox11_MouseEnter(object sender, EventArgs e)
        {
            pictureBox11.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox11.Cursor = Cursors.Hand;
        }

        private void pictureBox11_MouseLeave(object sender, EventArgs e)
        {
            pictureBox11.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox11.Cursor = Cursors.Default;
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {
            Form22 f = new Form22();
            f.Show();
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
