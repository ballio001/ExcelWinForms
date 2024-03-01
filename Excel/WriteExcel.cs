using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelWinForm.Excel
{
    class WriteExcel
    {
        public static void WriteExcelData(string[,] data, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //var to hold the objects
            Workbook wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            try
            {
                // Get the dimensions of the data array
                int rowCount = data.GetLength(0);
                int colCount = data.GetLength(1);

                // Write data to the worksheet
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        ws.Cells[i, j] = data[i - 1, j - 1];
                    }
                }

                // Save the workbook
                wb.SaveAs(filePath);
                wb.Close();

                MessageBox.Show("Data has been written to: " + filePath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        }
    }
}
