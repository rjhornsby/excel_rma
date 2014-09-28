using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using File = System.IO.File;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using AutoEllipsis;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {

        [DllImport("shlwapi.dll", CharSet = CharSet.Auto)]
        static extern bool PathCompactPathEx([Out] StringBuilder pszOut, string szPath, int cchMax, int dwFlags);

        private Excel.Application ExcelObj = null;
        private String billingFileName = "";
        private String responseFileName = "";
        private const int RESPONSE_ROW_OFFSET = 15;
        EllipsisFormat fmt = EllipsisFormat.None;

        public Form1()
        {
            InitializeComponent();
            ExcelObj = new Excel.Application();
            fmt |= EllipsisFormat.Start;
            fmt |= EllipsisFormat.Path;
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            Excel.Workbook billingWb = null;
            Excel.Workbook responseWb = null;
            lblBillingFile.ForeColor = SystemColors.ControlText;
            lblResponseFile.ForeColor = SystemColors.ControlText;

            if (billingFileName.Length == 0 || responseFileName.Length == 0)
            {
                if (billingFileName.Length == 0)
                {
                    lblBillingFile.ForeColor = Color.Red;
                }
                if (responseFileName.Length == 0)
                {
                    lblResponseFile.ForeColor = Color.Red;
                }
                return;
            }

            try
            {
                billingWb = ExcelObj.Workbooks.Open(billingFileName, Type.Missing, System.IO.FileAccess.Read);
            }
            catch (System.IO.FileNotFoundException ex)
            {
                MessageBox.Show(ex.Message, "File not found", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                   ex.Message, "Couldn't open file", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }


            try
            {
                responseWb = ExcelObj.Workbooks.Open(responseFileName);
            }
            catch (System.IO.FileNotFoundException ex)
            {
                MessageBox.Show(ex.Message, "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message, "Couldn't open file", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            Excel.Worksheet billingWs = billingWb.Worksheets[1];
            Excel.Worksheet responseWs = responseWb.Worksheets[1];

            Excel.Range billingUsedRange = billingWs.UsedRange;

            int totalRows = billingUsedRange.Rows.Count;

            foreach (Excel.Range billingRow in billingUsedRange.Rows) {
                // skip first row headers
                if (billingRow.Row == 1)
                    continue;

                Decimal progress = (Decimal)billingRow.Row / (Decimal)totalRows;
                progressBarClaims.Value = (int)Math.Round(progress * 100, 0);

                Excel.Range cell = (Excel.Range)billingRow.Columns["A"];
                textBoxClaimNum.Text = cell.Value2.ToString();

                // Get the responseRow - the row where we'll be copying TO
                int responseRowIdx = billingRow.Row + RESPONSE_ROW_OFFSET;
                Excel.Range responseRow = null;
                if ((Excel.Range)responseWs.Rows[responseRowIdx] == null)
                {
                    responseRow = (Excel.Range)responseWs.Rows[responseRowIdx];
                    responseRow.Insert();
                }
                else
                {
                    responseRow = (Excel.Range)responseWs.Rows[responseRowIdx];
                }

                copyColumns(billingRow, responseRow);

            }

            responseWb.Save();
            responseWb.Close();
            billingWb.Close();

            lblClaimNum.Text = "";
            System.Diagnostics.Debug.WriteLine("done!");
        }

        private void copyColumns(Excel.Range billingRow, Excel.Range responseRow)
        {
            this.copy(billingRow, "A", responseRow, "A");
            this.copy(billingRow, "I", responseRow, "D");
            this.copy(billingRow, "J", responseRow, "E");
            this.copy(billingRow, "M", responseRow, "H");
            this.copy(billingRow, "M", responseRow, "I");
            this.copy(billingRow, "O", responseRow, "J");
            this.copy(billingRow, "O", responseRow, "K");
            this.copy(billingRow, "Q", responseRow, "L");
            this.copy(billingRow, "Q", responseRow, "M");
            this.copy(billingRow, "S", responseRow, "N");
            this.copy(billingRow, "S", responseRow, "O");
            this.copy(billingRow, "AH", responseRow, "AD");
            this.copy(billingRow, "F", responseRow, "AE");
            this.copy(billingRow, "AJ", responseRow, "AK");
            this.copy(billingRow, "AR", responseRow, "AL");
            this.copy(billingRow, "AZ", responseRow, "AM");
            this.copy(billingRow, "BH", responseRow, "AN");
            this.copy(billingRow, "BP", responseRow, "AO");
            this.copy(billingRow, "BX", responseRow, "AP");
            this.copy(billingRow, "CF", responseRow, "AQ");
            this.copy(billingRow, "CX", responseRow, "AT");
            this.copy(billingRow, "X", responseRow, "AU");
        }

        private void copy(Excel.Range sourceRow, string sourceColumn, Excel.Range targetRow, string targetColumn)
        {
            Excel.Range cell = (Excel.Range)sourceRow.Columns[sourceColumn].Cells;
            if (cell.Value2 != null)
            {
                targetRow.Columns[targetColumn].Cells.Value2 = sourceRow.Columns[sourceColumn].Cells.Value2;
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void buttonPopBillingFileDialog_Click(object sender, EventArgs e)
        {
            lblBillingFile.ForeColor = SystemColors.ControlText;
            DialogResult result = openFileDialogBilling.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBoxBillingFileName.Text = this.billingFileName = openFileDialogBilling.FileName;
            }

        }

        private void buttonPopResponseFileDialog_Click(object sender, EventArgs e)
        {
            lblResponseFile.ForeColor = SystemColors.ControlText;
            DialogResult result = openFileDialogResponse.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBoxResponseFileName.Text = this.responseFileName = openFileDialogResponse.FileName;
            }
        }


        private void textBoxBillingFileName_MouseClick(object sender, MouseEventArgs e)
        {
            //buttonPopBillingFileDialog.PerformClick();
        }
        private void textBoxResponseFileName_MouseClick(object sender, MouseEventArgs e)
        {
            //buttonPopResponseFileDialog.PerformClick();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            for (int i = 1; (i <= 10); i++)
            {
                if ((worker.CancellationPending == true))
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    // Perform a time consuming operation and report progress.
                    System.Threading.Thread.Sleep(500);
                    worker.ReportProgress((i * 10));
                }
            }
        }



    }
}
