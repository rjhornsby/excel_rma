using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    class rma
    {
        private Excel.Application ExcelObj = null;
        private const int RESPONSE_ROW_OFFSET = 15;
        public String billingFileName { get; set;}
        public String responseFileName { get; set;}

        public rma()
        {
            ExcelObj = new Excel.Application();

        }

        public void process_rma() {
            Excel.Workbook billingWb = ExcelObj.Workbooks.Open(billingFileName);
            Excel.Workbook responseWb = ExcelObj.Workbooks.Open(responseFileName);
            System.Diagnostics.Debug.WriteLine(billingWb.Worksheets.Count);
            Excel.Worksheet billingWs = billingWb.Worksheets[1];
            Excel.Worksheet responseWs = responseWb.Worksheets[1];

            Excel.Range billingUsedRange = billingWs.UsedRange;

            int totalRows = billingUsedRange.Rows.Count;

            foreach (Excel.Range billingRow in billingUsedRange.Rows)
            {
                // skip first row headers
                if (billingRow.Row == 1)
                    continue;

                Decimal progress = (Decimal)billingRow.Row / (Decimal)totalRows;
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

                copyColumns(billingRow, responseRow);

            }

            responseWb.Save();
            responseWb.Close();
            billingWb.Close();
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

    }
}

