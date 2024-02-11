using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelWinForm.Excel
{
    class WriteExcel
    {
        public static void WriteExcelData(string[,] data)
        {
            //filepath to the location of the Excel
            string filePath = Files.OriginalFilePath;
            string filePathEdited = Files.EditedFilePath;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //var to hold the objects
            Workbook wb = excel.Workbooks.Open(filePath);
            Worksheet ws = wb.Worksheets[1];

            try
            {
                // Get the dimensions of the data array
                int rowCount = data.GetLength(0);
                int colCount = data.GetLength(1);

                // Get the range to write the data
                Microsoft.Office.Interop.Excel.Range range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCount, colCount]];

                // Write the data to the range
                range.Value = data;

                // Save the workbook and close Excel
                wb.SaveAs(filePathEdited);
                wb.Close();

                MessageBox.Show("Data has been written to: " + filePathEdited, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                wb.Close(false);
            }
            finally
            {
                excel.Quit();
            }
            // unocmment if you want to open excel after executing write
            //Process.Start(filePathEdited);
        }
    }
}
