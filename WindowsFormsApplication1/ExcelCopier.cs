using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTranscriptionMachine
{
    class ExcelCopier : IDisposable
    {
        private Excel.Application ExcelObj = null;
        public System.Windows.Forms.TextBox textBoxClaimNum { get; set; }
        public System.Windows.Forms.ProgressBar progressBarClaims { get; set; }
        public bool copySuccess = false;
        public String error = null;
        public bool isBusy = false;
        private const int RESPONSE_ROW_OFFSET = 15;
        private BackgroundWorker m_oWorker;

        public ExcelCopier(Form1 form)
        {
            ExcelObj = new Excel.Application();
            m_oWorker = new BackgroundWorker();
            // Create a background worker thread that ReportsProgress &
            // SupportsCancellation
            // Hook up the appropriate events.
            m_oWorker.DoWork += new DoWorkEventHandler(m_oWorker_DoWork);
            m_oWorker.ProgressChanged += new ProgressChangedEventHandler
                    (m_oWorker_ProgressChanged);
            m_oWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler
                    (m_oWorker_RunWorkerCompleted);
            m_oWorker.WorkerReportsProgress = true;
            m_oWorker.WorkerSupportsCancellation = true;
        }

        public void cancel()
        {
            if (m_oWorker.IsBusy)
            {
                m_oWorker.CancelAsync();
            }
        }

        void m_oWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            // This is probably a terrible way to pass in two worksheets...
            Excel.Worksheet[] wsArray = (Excel.Worksheet[])e.Argument;
            Excel.Worksheet billingWs = wsArray[0];
            Excel.Worksheet responseWs = wsArray[1];
            Excel.Range billingUsedRange = billingWs.UsedRange;
            int totalRows = billingUsedRange.Rows.Count;
        

            foreach (Excel.Range billingRow in billingUsedRange.Rows)
            {
                // skip first row headers
                if (billingRow.Row == 1)
                    continue;

                // Update the progress indicators
                int progress = (int)Math.Round(((Decimal)billingRow.Row / (Decimal)totalRows) * 100);
                Excel.Range cell = (Excel.Range)billingRow.Columns["A"];

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

                m_oWorker.ReportProgress(progress, cell.Value2.ToString());

                copyColumns(billingRow, responseRow);
                if (m_oWorker.CancellationPending)
                {
                    // Set the e.Cancel flag so that the WorkerCompleted event
                    // knows that the process was cancelled.
                    e.Cancel = true;
                    m_oWorker.ReportProgress(0, "");
                    return;
                }
            }

            //Report 100% completion on operation completed
            m_oWorker.ReportProgress(100, "");
        
            
        }

        void m_oWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBarClaims.Value = e.ProgressPercentage;
            this.textBoxClaimNum.Text = e.UserState.ToString();
        }

        void m_oWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.isBusy = false;
            if (e.Cancelled)
            {
                this.copySuccess = false;
                if (e.Error != null)
                {
                    this.error = e.Error.Message;
                }
            }
            else
            {
                this.copySuccess = true;
            }
        }

        public void doCopy(Excel.Worksheet billingWs, Excel.Worksheet responseWs)
        {
            this.isBusy = true;
            Excel.Worksheet[] wsArray = { billingWs, responseWs };
            m_oWorker.RunWorkerAsync(wsArray);

            /*
            Excel.Range billingUsedRange = billingWs.UsedRange;

            int totalRows = billingUsedRange.Rows.Count;

            foreach (Excel.Range billingRow in billingUsedRange.Rows)
            {
                // skip first row headers
                if (billingRow.Row == 1)
                    continue;

                // Update the progress indicators
                Decimal progress = (Decimal)billingRow.Row / (Decimal)totalRows;
                Excel.Range cell = (Excel.Range)billingRow.Columns["A"];
                this.progressBarClaims.Value = (int)Math.Round(progress * 100, 0);
                this.textBoxClaimNum.Text = cell.Value2.ToString();


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

            */

        }

        private void copyColumns(Excel.Range billingRow, Excel.Range responseRow)
        {
            this.copyCell(billingRow, "A", responseRow, "A");
            this.copyCell(billingRow, "I", responseRow, "D");
            this.copyCell(billingRow, "J", responseRow, "E");
            this.copyCell(billingRow, "M", responseRow, "H");
            this.copyCell(billingRow, "M", responseRow, "I");
            this.copyCell(billingRow, "O", responseRow, "J");
            this.copyCell(billingRow, "O", responseRow, "K");
            this.copyCell(billingRow, "Q", responseRow, "L");
            this.copyCell(billingRow, "Q", responseRow, "M");
            this.copyCell(billingRow, "S", responseRow, "N");
            this.copyCell(billingRow, "S", responseRow, "O");
            this.copyCell(billingRow, "AH", responseRow, "AE");
            this.copyCell(billingRow, "F", responseRow, "AF");
            this.copyCell(billingRow, "AJ", responseRow, "AL");
            this.copyCell(billingRow, "AR", responseRow, "AM");
            this.copyCell(billingRow, "AZ", responseRow, "AN");
            this.copyCell(billingRow, "BH", responseRow, "AO");
            this.copyCell(billingRow, "BP", responseRow, "AP");
            this.copyCell(billingRow, "BX", responseRow, "AQ");
            this.copyCell(billingRow, "CF", responseRow, "AR");
            this.copyCell(billingRow, "CX", responseRow, "AU");
            this.copyCell(billingRow, "X", responseRow, "AV");
        }

        private void copyCell(Excel.Range sourceRow, string sourceColumn, Excel.Range targetRow, string targetColumn)
        {
            Excel.Range sourceColumnRange = (Excel.Range)sourceRow.Columns[sourceColumn];
            Excel.Range sourceCell = sourceColumnRange.Cells;

            Excel.Range targetColumnRange = (Excel.Range)targetRow.Columns[targetColumn];
            Excel.Range targetCell = targetColumnRange.Cells;

            if (sourceCell.Value2 != null)
            {
                targetCell.Value2 = sourceCell.Value2;
            }
        }

        public void Dispose() {
            m_oWorker.Dispose();
        }

    }
}

