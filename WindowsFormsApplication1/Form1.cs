using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using File = System.IO.File;
using Excel = Microsoft.Office.Interop.Excel;
using AutoEllipsis;

namespace ExcelTranscriptionMachine
{
    public partial class Form1 : Form
    {

        private Excel.Application ExcelObj = null;
        private String billingFileName = "";
        private String responseFileName = "";
        private const int RESPONSE_ROW_OFFSET = 15;
        EllipsisFormat fmt = EllipsisFormat.None;
        ExcelCopier copier = null;

        public Form1()
        {
            InitializeComponent();
            ExcelObj = new Excel.Application();
            fmt |= EllipsisFormat.Start;
            fmt |= EllipsisFormat.Path;

        }

        private void enableControls()
        {
            btnGo.Enabled = true;
            buttonPopBillingFileDialog.Enabled = true;
            buttonPopResponseFileDialog.Enabled = true;
            textBoxBillingFileName.Enabled = true;
            textBoxResponseFileName.Enabled = true;
        }
        private void disableControls()
        {
            btnGo.Enabled = false;
            buttonPopBillingFileDialog.Enabled = false;
            buttonPopResponseFileDialog.Enabled = false;
            textBoxBillingFileName.Enabled = false;
            textBoxResponseFileName.Enabled = false;
        }

        private void btnGo_Click(object sender, EventArgs e)
        {

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

            if (textBoxBillingFileName.FullText == textBoxResponseFileName.FullText)
            {
                MessageBox.Show("Billing file must be different from response file.", "Operation not possible");
                return;
            }

            this.disableControls();
            textBoxClaimNum.Text = "Copier warming up, please wait...";

            Excel.Workbook billingWb = null;
            Excel.Workbook responseWb = null;
            lblBillingFile.ForeColor = SystemColors.ControlText;
            lblResponseFile.ForeColor = SystemColors.ControlText;

            try
            {
                billingWb = ExcelObj.Workbooks.Open(billingFileName, Type.Missing, System.IO.FileAccess.Read);
            }
            catch (System.IO.FileNotFoundException ex)
            {
                MessageBox.Show(ex.Message, "File not found", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                this.enableControls();
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                   ex.Message, "Couldn't open file", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                this.enableControls();
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

            copier = new ExcelCopier(this);
            copier.progressBarClaims = this.progressBarClaims;
            copier.textBoxClaimNum = this.textBoxClaimNum;

            btnCancel.Enabled = true;
            copier.doCopy((Excel.Worksheet)billingWb.Sheets.get_Item(1), (Excel.Worksheet)responseWb.Sheets.get_Item(1));

            while (copier.isBusy)
            {
                Application.DoEvents();
            }
            if (copier.copySuccess) {
                try
                {
                    responseWb.Save();
                    responseWb.Close(true);
                    MessageBox.Show("Excel copier job complete");
                }
                catch (System.IO.IOException ex)
                {
                    MessageBox.Show(ex.Message, "Unable to save response file");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Unable to save response file");
                }

            }
            else
            {
                responseWb.Close(false);
                MessageBox.Show("Excel copier job canceled.");
                if (copier.error != null)
                {
                    MessageBox.Show(copier.error);
                }
                Random random = new Random();
                if (random.Next(100) >= 95)
                {
                    textBoxClaimNum.Text = "PC LOAD LETTER";
                }
            }
            
            billingWb.Close(false);

            enableControls();
            btnCancel.Enabled = false;

            textBoxClaimNum.Text = "";

            copier = null;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.copier.cancel();
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

        private void textBoxBillingFileName_TextChanged(object sender, EventArgs e)
        {
            this.billingFileName = openFileDialogBilling.FileName = textBoxBillingFileName.FullText;
        }

        private void textBoxResponseFileName_TextChanged(object sender, EventArgs e)
        {
            this.responseFileName = openFileDialogResponse.FileName = textBoxResponseFileName.FullText;
        }
    }
}
