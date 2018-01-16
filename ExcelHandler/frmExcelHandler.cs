using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace ExcelHandler
{
    public partial class frmExcelHandler : Form
    {
        public frmExcelHandler()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Choose file";
            ofd.Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx";
            ofd.Multiselect = false;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                btnRun.Enabled = false;
                float limitAmount = float.Parse(textBox1.Text);
                string text = txtCustomText.Text;
                bgdWorker.RunWorkerAsync(new object[] { ofd.FileName, limitAmount, text });
            }
        }

        private float Str2Float(string input)
        {
            float result = 0;
            string resultingStr = "";
            resultingStr = string.Join(string.Empty, Regex.Matches(input, @"^\-?[0-9]+(?:\.[0-9]+)?$").OfType<Match>().Select(m => m.Value));
            float.TryParse(resultingStr, out result);
            return result;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //only allow integer (no decimal point)
            if (!char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-') && (e.KeyChar != 8))
                e.Handled = true;
            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
                e.Handled = true;
            // only allow sign symbol at first char
            if ((e.KeyChar == '-') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('-') > -1))
                e.Handled = true;
            if ((e.KeyChar == '-') && !((sender as System.Windows.Forms.TextBox).Text.IndexOf('-') > -1) && ((sender as System.Windows.Forms.TextBox).SelectionStart != 0))
                e.Handled = true;
        }

        private void bgdWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string fileName = (e.Argument as object[])[0].ToString();
            float limitAmount = float.Parse((e.Argument as object[])[1].ToString());
            string text = (e.Argument as object[])[2].ToString();
            StringBuilder sb = new StringBuilder();
            bgdWorker.ReportProgress(1, "Opening file " + System.IO.Path.GetFileName(fileName) + "...");
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(fileName);
            try
            {
                foreach (Worksheet sheet in workbook.Sheets)
                {
                    bgdWorker.ReportProgress(1, "Working with worksheet " + sheet.Name + "...");
                    int counter = 7;
                    float sum = 0;
                    Range cell;
                    //sum
                    while (true)
                    {
                        cell = sheet.Cells[counter, 2];
                        if (cell.Value == null)
                            break;
                        sum += Str2Float(((object)cell.Value).ToString());
                        counter++;
                    }
                    if (sum >= limitAmount || (sheet.Cells[1, 2].Value as object).ToString().Contains("會員"))
                    {
                        sb.Append((sheet.Cells[1, 2].Value as object).ToString().Split(new string[] { "\n" }, StringSplitOptions.None)[0]);
                        //get phone number
                        foreach(string line in (sheet.Cells[1, 5].Value as object).ToString().Split(new string[] { "\n" }, StringSplitOptions.None))
                        {
                            string temp = string.Join(string.Empty, Regex.Matches(line, @"\d+").OfType<Match>().Select(m => m.Value));
                            if (Regex.Match(temp, @"^(09[0-9]{8})$").Success)
                                sb.Append(", phone: " + temp);
                        }
                        sb.AppendLine();
                        //write to excel
                        if (chkWrite.Checked)
                        {
                            while (true)
                            {
                                cell = sheet.Cells[counter, 4];
                                if (cell.Value == null)
                                {
                                    cell.Value = text;
                                    cell.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                    break;
                                }
                                counter++;
                            }
                        }
                    }
                    if (chkWrite.Checked)
                        workbook.Save();
                }
            }
            catch (Exception)
            { }
            finally
            {
                workbook.Close(false);
                Marshal.FinalReleaseComObject(workbook);
                workbook = null;
                excel.Quit();
                Marshal.FinalReleaseComObject(excel);
                excel = null;
            }
            bgdWorker.ReportProgress(100, sb.ToString());
        }

        private void bgdWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Clipboard.SetText(txtResult.Text);
            MessageBox.Show("Done, result string copied to clipboard");
            btnRun.Enabled = true;
            lblStatus.Text = "Idle";
        }

        private void bgdWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 100)
                txtResult.Text = e.UserState.ToString();
            else
                lblStatus.Text = e.UserState.ToString();
        }
    }
}
