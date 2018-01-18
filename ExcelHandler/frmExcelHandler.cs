using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using System.Threading;

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
                float limitAmount = float.Parse(txtLimitAmount.Text);
                string text = txtCustomText.Text;
                bgdWorker.RunWorkerAsync(new object[] { ofd.FileName, limitAmount, text });
            }
        }

        private float Str2Float(string input)
        {
            float result = 0;
            string resultingStr = "";
            input = input.Replace(",", "");
            input = input.Replace("$", "");
            resultingStr = string.Join(string.Empty, Regex.Matches(input, @"^\-?[0-9]+(?:\.[0-9]+)?$").OfType<Match>().Select(m => m.Value));
            float.TryParse(resultingStr, out result);
            return result;
        }

        private void txtLimitAmount_KeyPress(object sender, KeyPressEventArgs e)
        {
            //only allow integer (no decimal point)
            if (!char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-') && (e.KeyChar != 8))
                e.Handled = true;
            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
                e.Handled = true;
            // only allow sign symbol at first char
            if ((e.KeyChar == '-') && ((sender as TextBox).Text.IndexOf('-') > -1))
                e.Handled = true;
            if ((e.KeyChar == '-') && !((sender as TextBox).Text.IndexOf('-') > -1) && ((sender as TextBox).SelectionStart != 0))
                e.Handled = true;
        }

        private void bgdWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string fileName = (e.Argument as object[])[0].ToString();
            float limitAmount = float.Parse((e.Argument as object[])[1].ToString());
            string text = (e.Argument as object[])[2].ToString();
            StringBuilder sb = new StringBuilder();
            IWorkbook workbook = null;
            try
            {
                bgdWorker.ReportProgress(1, "Opening file " + Path.GetFileName(fileName) + "...");
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    if (Path.GetExtension(fileName).Equals(".xls"))
                        workbook = new HSSFWorkbook(fs);
                    else if (Path.GetExtension(fileName).Equals(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        float sum = 0;
                        ISheet sheet = workbook.GetSheetAt(i);
                        bgdWorker.ReportProgress(1, "Working with worksheet " + sheet.SheetName + "...");
                        //sum
                        for (int j = 6; j < sheet.LastRowNum; j++)
                            sum += Str2Float(sheet.GetRow(j).GetCell(1).ToString());
                        //determine
                        string name = sheet.GetRow(0).GetCell(1).ToString();
                        if (sum >= limitAmount || (name.Contains("會員") && !name.Contains("非會員")))
                        {
                            sb.Append(name.Split(new string[] { "\n" }, StringSplitOptions.None)[0]);
                            //get phone number
                            foreach (string line in sheet.GetRow(0).GetCell(4).ToString().Split(new string[] { "\n" }, StringSplitOptions.None))
                            {
                                string temp = string.Join(string.Empty, Regex.Matches(line, @"\d+").OfType<Match>().Select(m => m.Value));
                                if (Regex.Match(temp, @"^(09[0-9]{8})$").Success)
                                    sb.Append(", phone: " + temp);
                            }
                            sb.AppendLine();
                            //write to excel
                            if (chkWrite.Checked)
                            {
                                for (int j = 6; j < sheet.LastRowNum; j++)
                                {
                                    if (sheet.GetRow(j).GetCell(3).ToString().Equals(""))
                                    {
                                        sheet.GetRow(j).GetCell(3).SetCellValue(text);
                                        sheet.GetRow(j).GetCell(3).SetCellType(CellType.String);
                                        ICellStyle styleLeft = workbook.CreateCellStyle();
                                        styleLeft.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                                        styleLeft.VerticalAlignment = VerticalAlignment.Center;
                                        styleLeft.WrapText = true; //wrap the text in the cell
                                        sheet.GetRow(j).GetCell(3).CellStyle = styleLeft;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    bgdWorker.ReportProgress(100, sb.ToString());
                }
                if (chkWrite.Checked)
                {
                    Thread SaveFileThread = new Thread(new ParameterizedThreadStart(SaveFile));
                    SaveFileThread.SetApartmentState(ApartmentState.STA);
                    SaveFileThread.Start(workbook);
                    SaveFileThread.Join();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.StackTrace + ": " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void SaveFile(object workbook)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Choose output file name";
            sfd.Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.ReadWrite))
                        ((IWorkbook)workbook).Write(fs);
                    MessageBox.Show("Save file success");
                }
                catch(Exception)
                { }
            }
        }

        private void bgdWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!txtResult.Text.Equals(""))
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
