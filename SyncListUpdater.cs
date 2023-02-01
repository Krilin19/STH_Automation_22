using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using STH_Automation_22;

namespace STH_Automation_22
{
    
    public partial class SyncListUpdater : Form
    {
        
        public SyncListUpdater()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            string Sync_Manager = @"T:\Lopez\Sync_Manager.xlsx";
            SyncListUpdater SyncListUpdater_ = new SyncListUpdater();
            
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Excel Sync Manager file not found, try Sync the normal way", "Sync Warning");
                MessageBox.Show("Another user is using the Excel Sync Manager file, Try again", "Sync Warning");
                //return Autodesk.Revit.UI.Result.Cancelled;
            }
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
                var Time_ = DateTime.Now;

                for (int row = 1; row < 20; row++)
                {
                    if (sheet.Cells[row, 1].Value == null)
                    {
                        break;
                    }

                    if (sheet.Cells[row, 1].Value != null)
                    {
                        var Value1 = sheet.Cells[row, 1].Value;
                        var Value2 = sheet.Cells[row, 2].Value;
                        //s += Value1 + " + " + Value2.ToString() + "\n";
                        SyncListUpdater_.listBox1.Items.Add(Value1 + " + " + Value2.ToString() + "\n");

                    }

                }
                package.Save();
            }
            
            SyncListUpdater_.listBox1.Update();
            SyncListUpdater_.listBox1.Refresh();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            Start_();
            this.DialogResult = DialogResult.OK;
        }

        public bool Start_()
        {
            bool timerToCheck = true;

            return timerToCheck;
        }
    }
}
