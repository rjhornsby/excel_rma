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

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {

        private Excel.Application ExcelObj = null;
        private String billingFileName = "";
        private String responseFileName = "";
        private const int RESPONSE_ROW_OFFSET = 15;
        EllipsisFormat fmt = EllipsisFormat.None;
        BackgroundWorker m_oWorker;

        public Form1()
        {
            InitializeComponent();
            ExcelObj = new Excel.Application();
            fmt |= EllipsisFormat.Start;
            fmt |= EllipsisFormat.Path;
            // Create a background worker thread that ReportsProgress &
            // SupportsCancellation
            // Hook up the appropriate events.
            /*m_oWorker.DoWork += new DoWorkEventHandler(ExcelCopier.m_oWorker_DoWork);
            m_oWorker.ProgressChanged += new ProgressChangedEventHandler
                    (m_oWorker_ProgressChanged);
            m_oWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler
                    (m_oWorker_RunWorkerCompleted);
            m_oWorker.WorkerReportsProgress = true;
            m_oWorker.WorkerSupportsCancellation = true;
        */
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

            ExcelCopier copier = new ExcelCopier(this);
            copier.progressBarClaims = this.progressBarClaims;
            copier.textBoxClaimNum = this.textBoxClaimNum;

            copier.doCopy(billingWb, responseWb);

            responseWb.Save();
            responseWb.Close();
            billingWb.Close();
           
            lblClaimNum.Text = "";
            System.Diagnostics.Debug.WriteLine("done!");
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

        private void btnStartAsyncOperation_Click(object sender, EventArgs e)
        {
            //Change the status of the buttons on the UI accordingly
            //The start button is disabled as soon as the background operation is started
            //The Cancel button is enabled so that the user can stop the operation 
            //at any point of time during the execution
            btnGo.Enabled = false;
            btnCancel.Enabled = true;

            // Kickoff the worker thread to begin it's DoWork function.
            m_oWorker.RunWorkerAsync();
        }


        private void btnCancel_Click(object sender, EventArgs e)
        {

            if (m_oWorker.IsBusy)
            {

                // Notify the worker thread that a cancel has been requested.
                // The cancel will not actually happen until the thread in the
                // DoWork checks the m_oWorker.CancellationPending flag. 
                m_oWorker.CancelAsync();
            }
        }

    }
}
