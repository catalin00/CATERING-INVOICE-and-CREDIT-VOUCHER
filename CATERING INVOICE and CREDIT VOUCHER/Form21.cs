using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;



namespace CATERING_INVOICE_and_CREDIT_VOUCHER
{
    public partial class Form21 : Form
    {
        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb; Persist Security Info=False;");
        string bckpath;
        string restorepath;
        public Form21()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            DialogResult rez = folder.ShowDialog();
            if (rez==DialogResult.OK)
            {
                bckpath = folder.SelectedPath;
                textBox1.Text = bckpath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            DialogResult rez = file.ShowDialog();
            if (rez==DialogResult.OK)
            {
                restorepath = file.FileName;
                textBox2.Text = restorepath;

            }

        }
        void savexml()
        {
            if (connection.State != ConnectionState.Open) connection.Open();
            object misValue = System.Reflection.Missing.Value;
            string data = null;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Excel.Worksheet xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet1.Name = "Customers";
            OleDbDataAdapter dscmd = new OleDbDataAdapter("select * from Customers", connection);
            DataSet ds = new DataSet();
            dscmd.Fill(ds);
            xlWorkSheet1.Cells[1, 1] = "CustomerID";
            xlWorkSheet1.Cells[1, 2] = "Salutation";
            xlWorkSheet1.Cells[1, 3] = "FirstName";
            xlWorkSheet1.Cells[1, 4] = "LastName";
            xlWorkSheet1.Cells[1, 5] = "Title";
            xlWorkSheet1.Cells[1, 6] = "MobileNumber";
            xlWorkSheet1.Cells[1, 7] = "Departmant";
            xlWorkSheet1.Cells[1, 8] = "CompanyName";
            xlWorkSheet1.Cells[1, 9] = "PhoneNumber";
            xlWorkSheet1.Cells[1, 10] = "PhoneExt";
            xlWorkSheet1.Cells[1, 11] = "FaxNumber";
            xlWorkSheet1.Cells[1, 12] = "Email";
            xlWorkSheet1.Cells[1, 13] = "Website";
            xlWorkSheet1.Cells[1, 14] = "Address";
            xlWorkSheet1.Cells[1, 15] = "City";
            xlWorkSheet1.Cells[1, 16] = "State";
            xlWorkSheet1.Cells[1, 17] = "PostalCode";
            xlWorkSheet1.Cells[1, 18] = "CateringAddress";
            xlWorkSheet1.Cells[1, 19] = "CateringCity";
            xlWorkSheet1.Cells[1, 20] = "CateringState";
            xlWorkSheet1.Cells[1, 21] = "CateringZIP";
            xlWorkSheet1.Cells[1, 22] = "Notes";

            for (int i=0; i<ds.Tables[0].Rows.Count; i++)
            {
                for (int j=0; j < ds.Tables[0].Columns.Count; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet1.Cells[i + 2, j + 1].NumberFormat = "@";
                    xlWorkSheet1.Cells[i + 2, j + 1] = data;

                }
            }


            ds.Dispose();
            dscmd.Dispose();
            
            Excel.Worksheet xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlWorkSheet2.Name = "FoodProducts";
             dscmd = new OleDbDataAdapter("select * from FoodProducts", connection);
             ds = new DataSet();
            dscmd.Fill(ds);
            xlWorkSheet2.Cells[1, 1] = "ProductID";
            xlWorkSheet2.Cells[1, 2] = "DishName";
            xlWorkSheet2.Cells[1, 3] = "DishDescription";
            xlWorkSheet2.Cells[1, 4] = "TaxableFood";
            xlWorkSheet2.Cells[1, 5] = "UnitPrice";
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet2.Cells[i + 2, j + 1].NumberFormat = "@";
                    xlWorkSheet2.Cells[i + 2, j + 1] = data;

                }
            }



            ds.Dispose();
            dscmd.Dispose();
            
            Excel.Worksheet xlWorkSheet3 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlWorkSheet3.Name = "VoucherLog";
            dscmd = new OleDbDataAdapter("select * from VoucherLog", connection);
            ds = new DataSet();
            dscmd.Fill(ds);
            xlWorkSheet3.Cells[1, 1] = "VoucherID";
            xlWorkSheet3.Cells[1, 2] = "CustomerID";
            xlWorkSheet3.Cells[1, 3] = "VoucherDate";
            xlWorkSheet3.Cells[1, 4] = "CreditItem";
            xlWorkSheet3.Cells[1, 5] = "Reason";
            xlWorkSheet3.Cells[1, 6] = "EmployeeIDreq";
            xlWorkSheet3.Cells[1, 7] = "ManagerApprv";
            xlWorkSheet3.Cells[1, 8] = "DateApprv";
            xlWorkSheet3.Cells[1, 9] = "EmployeeIDApply";
            xlWorkSheet3.Cells[1, 10] = "DateApplied";
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet3.Cells[i + 2, j + 1].NumberFormat = "@";
                    xlWorkSheet3.Cells[i + 2, j + 1] = data;

                }
            }

            ds.Dispose();
            dscmd.Dispose();
            
            Excel.Worksheet xlWorkSheet4 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlWorkSheet4.Name = "InvoiceHeader";
            dscmd = new OleDbDataAdapter("select * from InvoiceHeader", connection);
            ds = new DataSet();
            dscmd.Fill(ds);
            xlWorkSheet4.Cells[1, 1] = "InvoiceID";
            xlWorkSheet4.Cells[1, 2] = "CustomerID";
            xlWorkSheet4.Cells[1, 3] = "EmployeeID";
            xlWorkSheet4.Cells[1, 4] = "DeliveryID";
            xlWorkSheet4.Cells[1, 5] = "InvoiceDate";
            xlWorkSheet4.Cells[1, 6] = "CaterDate";
            xlWorkSheet4.Cells[1, 7] = "CaterTime";
            xlWorkSheet4.Cells[1, 8] = "OrderType";
            xlWorkSheet4.Cells[1, 9] = "Racks";
            xlWorkSheet4.Cells[1, 10] = "Sterno";
            xlWorkSheet4.Cells[1, 11] = "Subtotal";
            xlWorkSheet4.Cells[1, 12] = "DiscountApprv";
            xlWorkSheet4.Cells[1, 13] = "Discount";
            xlWorkSheet4.Cells[1, 14] = "SalesTax";
            xlWorkSheet4.Cells[1, 15] = "RackDeposit";
            xlWorkSheet4.Cells[1, 16] = "RackReturned";
            xlWorkSheet4.Cells[1, 17] = "RackRetDate";
            xlWorkSheet4.Cells[1, 18] = "Total";
            xlWorkSheet4.Cells[1, 19] = "Payments";
            xlWorkSheet4.Cells[1, 20] = "Balance";
            xlWorkSheet4.Cells[1, 21] = "Notes";
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet4.Cells[i + 2, j + 1].NumberFormat = "@";
                    xlWorkSheet4.Cells[i + 2, j + 1] = data;

                }
            }




            ds.Dispose();
            dscmd.Dispose();
            
            Excel.Worksheet xlWorkSheet5 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlWorkSheet5.Name = "InvoiceDetails";
            dscmd = new OleDbDataAdapter("select * from InvoiceDetails", connection);
            ds = new DataSet();
            dscmd.Fill(ds);
            xlWorkSheet5.Cells[1, 1] = "ID";
            xlWorkSheet5.Cells[1, 2] = "InvoiceID";
            xlWorkSheet5.Cells[1, 3] = "ProductID";
            xlWorkSheet5.Cells[1, 4] = "DishName";
            xlWorkSheet5.Cells[1, 5] = "DishDescription";
            xlWorkSheet5.Cells[1, 6] = "Quantity";
            xlWorkSheet5.Cells[1, 7] = "UnitPrice";
            xlWorkSheet5.Cells[1, 8] = "Taxable";
            xlWorkSheet5.Cells[1, 9] = "Notes";
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet5.Cells[i + 2, j + 1].NumberFormat = "@";
                    xlWorkSheet5.Cells[i + 2, j + 1] = data;

                }
            }



            ds.Dispose();
            dscmd.Dispose();
            
            Excel.Worksheet xlWorkSheet6 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlWorkSheet6.Name = "Payments";
            dscmd = new OleDbDataAdapter("select * from Payments", connection);
            ds = new DataSet();
            dscmd.Fill(ds);
            xlWorkSheet6.Cells[1, 1] = "PaymentID";
            xlWorkSheet6.Cells[1, 2] = "InvoiceID";
            xlWorkSheet6.Cells[1, 3] = "CustomerID";
            xlWorkSheet6.Cells[1, 4] = "PaymentDate";
            xlWorkSheet6.Cells[1, 5] = "PaymentAmount";
            xlWorkSheet6.Cells[1, 6] = "PaymentType";
            xlWorkSheet6.Cells[1, 7] = "MicrosRefNo";

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet6.Cells[i + 2, j + 1].NumberFormat = "@";
                    xlWorkSheet6.Cells[i + 2, j + 1] = data;

                }
            }



            ds.Dispose();
            dscmd.Dispose();
            
            Excel.Worksheet xlWorkSheet7 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlWorkSheet7.Name = "Settings";
            dscmd = new OleDbDataAdapter("select * from Settings", connection);
            ds = new DataSet();
            dscmd.Fill(ds);
            xlWorkSheet7.Cells[1, 1] = "ID";
            xlWorkSheet7.Cells[1, 2] = "LogoFile";
            xlWorkSheet7.Cells[1, 3] = "Logo";
            xlWorkSheet7.Cells[1, 4] = "CompanyName";
            xlWorkSheet7.Cells[1, 5] = "Address";
            xlWorkSheet7.Cells[1, 6] = "City";
            xlWorkSheet7.Cells[1, 7] = "State";
            xlWorkSheet7.Cells[1, 8] = "ZipCode";
            xlWorkSheet7.Cells[1, 9] = "Phone";
            xlWorkSheet7.Cells[1, 10] = "Fax";
            xlWorkSheet7.Cells[1, 11] = "URL";
            xlWorkSheet7.Cells[1, 12] = "eMail";
            xlWorkSheet7.Cells[1, 13] = "TaxRate";
            xlWorkSheet7.Cells[1, 14] = "RackCharge";
            xlWorkSheet7.Cells[1, 15] = "Paidico";
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    //xlWorkSheet7.Cells[i + 1, j + 1].NumberFormat = "@";
                    xlWorkSheet7.Cells[i + 2, j + 1] = data;

                }
            }

            xlWorkBook.SaveAs(bckpath + "/umbertos_backup_" + DateTime.Today.ToString("MM.dd.yyyy")+".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();


            releaseObject(xlWorkSheet1);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);




        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (bckpath != "" && bckpath != null)
            {
                if (checkBox2.Checked)
                {
                    File.Copy(AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb", bckpath + "/umbertos_backup_" + DateTime.Today.ToString("MM.dd.yyyy") + ".accdb", true);
                    
                    
                }
                if (checkBox1.Checked)
                {
                    savexml();
                    
                }
                if (!checkBox1.Checked && !checkBox2.Checked) MessageBox.Show("Select Backup Type");
                else MessageBox.Show("Backup Complete");
            }
            else MessageBox.Show("select destination folder");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (restorepath != "" && restorepath != null)
            {
                File.Copy(restorepath, AppDomain.CurrentDomain.BaseDirectory + "umbertos.accdb", true);
                MessageBox.Show("Backup Restored!");
            }
            else MessageBox.Show("select a database");
        }
    }
}
