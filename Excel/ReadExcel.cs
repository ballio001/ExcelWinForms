using Microsoft.Office.Interop.Excel;
using System.Text;

namespace ExcelWinForm.Excel
{
    class ReadExcel
    {
        public static string ReadExcelData()
        {
            string filePath = Files.EditedFilePath;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            StringBuilder data = new StringBuilder();

            //var to hold the objects
            Workbook wb;
            Worksheet ws;

            //opens workbook and stores in wb, same with sheet
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];

            // Get the used range of the worksheet
            Range usedRange = ws.UsedRange;

            // Get the number of rows and columns in the used range
            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            // Iterate through all rows and columns to read the data
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    // Get the cell value
                    Range cell = ws.Cells[i, j];
                    string cellValue = cell.Value != null ? cell.Value.ToString() : "";

                    // Append the cell value to the data string
                    data.Append(cellValue + "\t");
                }
                // Add a new line after each row
                data.AppendLine();
            }

            // Close the workbook and Excel application
            wb.Close();
            excel.Quit();

            // Return the collected data
            return data.ToString();
        }
    }
}
