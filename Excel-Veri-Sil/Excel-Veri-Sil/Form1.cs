using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Drawing.Printing;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;


namespace Excel_Veri_Sil
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        string dosya_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Stok\\----dosyaadı----;Extended Properties='Excel 8.0;HDR=YES';";
        string dosya_yolu2 = "C:\\Stok\\dosyaadı.xlsx";
        private void delete_fonk(int sayfa, string s2)
        {
            
            OleDbConnection conn_up = new OleDbConnection(dosya_yolu);
            conn_up.Open();
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter("SELECT * FROM [Sayfa" + s2 + "$] ", conn_up);
            DataTable dt = new DataTable();
            dataAdapter.Fill(dt);
            conn_up.Close();

            GC.GetTotalMemory(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.GetTotalMemory(true);


            //System.Threading.Thread.Sleep(3000);

            String expression = dt.Columns[0].ToString() + " = '" + textBox1.Text + "'";

            DataRow[] dr = dt.Select(expression);
            if (0 != dr.Length)
            {
                Excel.Application ExcelApp = new Excel.Application();
                //System.Threading.Thread.Sleep(3000);

                ExcelApp.Visible = false;

                Excel.Workbook ExcelWorkbook = ExcelApp.Workbooks.Open(dosya_yolu2);
                Excel.Worksheet ExcelWorksheet = ExcelWorkbook.Sheets[sayfa];
                //System.Threading.Thread.Sleep(3000);

                for (int index = 0; index < dr.Length; index++)
                {
                    string toDelete = "A" + (dt.Rows.IndexOf(dr[index]) + 2).ToString();
                    Excel.Range cells = (Excel.Range)ExcelWorksheet.Range[toDelete, Type.Missing];
                    cells.EntireRow.Delete();

                }

                ExcelWorkbook.Save();
                ExcelWorkbook.Close();
                ExcelApp.Quit();
                ExcelApp.Quit();

                Marshal.ReleaseComObject(ExcelWorksheet);
                Marshal.ReleaseComObject(ExcelWorksheet);


                Marshal.ReleaseComObject(ExcelWorkbook);
                Marshal.ReleaseComObject(ExcelWorkbook);
                Marshal.ReleaseComObject(ExcelApp);
                Marshal.ReleaseComObject(ExcelApp);


                ExcelWorksheet = null;
                ExcelWorksheet = null;

                ExcelWorkbook = null;
                ExcelWorkbook = null;
                ExcelApp = null;

                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);

            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Value = 25;
            delete_fonk(1, "");
            GC.GetTotalMemory(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.GetTotalMemory(true);
            progressBar1.Value = 50;

            delete_fonk(2, "1");
            GC.GetTotalMemory(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.GetTotalMemory(true);
            progressBar1.Value = 75;
            delete_fonk(3, "2");
            GC.GetTotalMemory(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.GetTotalMemory(true);


            progressBar1.Value = 100;
            System.Threading.Thread.Sleep(2000);
            MessageBox.Show("" + textBox1.Text + " Numaralı Ürün Silinmiştir.", "Bilgi Penceresi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            progressBar1.Visible = false;

            textBox1.Clear();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            

        }
    }
}
