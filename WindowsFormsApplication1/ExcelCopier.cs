using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    class ExcelCopier
    {
        private Excel.Application ExcelObj = null;
        public System.Windows.Forms.TextBox textBoxClaimNum { get; set; }
        public System.Windows.Forms.ProgressBar progressBarClaims { get; set; }
        private const int RESPONSE_ROW_OFFSET = 15;

        public ExcelCopier(Form1 form)
        {
            ExcelObj = new Excel.Application();
        }

        public void m_oWorker_DoWork(object sender, DoWorkEventArgs e)
        { }

        public void doCopy(Excel.Workbook billingWb, Excel.Workbook responseWb)
        {
            Excel.Worksheet billingWs = billingWb.Worksheets[1];
            Excel.Worksheet responseWs = responseWb.Worksheets[1];

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
            this.copyCell(billingRow, "AH", responseRow, "AD");
            this.copyCell(billingRow, "F", responseRow, "AE");
            this.copyCell(billingRow, "AJ", responseRow, "AK");
            this.copyCell(billingRow, "AR", responseRow, "AL");
            this.copyCell(billingRow, "AZ", responseRow, "AM");
            this.copyCell(billingRow, "BH", responseRow, "AN");
            this.copyCell(billingRow, "BP", responseRow, "AO");
            this.copyCell(billingRow, "BX", responseRow, "AP");
            this.copyCell(billingRow, "CF", responseRow, "AQ");
            this.copyCell(billingRow, "CX", responseRow, "AT");
            this.copyCell(billingRow, "X", responseRow, "AU");
        }

        private void copyCell(Excel.Range sourceRow, string sourceColumn, Excel.Range targetRow, string targetColumn)
        {
            Excel.Range cell = (Excel.Range)sourceRow.Columns[sourceColumn].Cells;
            if (cell.Value2 != null)
            {
                targetRow.Columns[targetColumn].Cells.Value2 = sourceRow.Columns[sourceColumn].Cells.Value2;
            }
        }


    }
}

