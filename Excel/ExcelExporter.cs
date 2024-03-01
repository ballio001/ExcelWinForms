using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWinForm.Excel
{
    public class ExcelExport : IDataExporter
    {
        private string filePath;

        public ExcelExport(string filePath)
        {
            this.filePath = filePath;
        }

        public void ExportData(List<Person> persons)
        {
            // Create Excel application instance
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                // Create a new workbook
                Workbook wb = excel.Workbooks.Add();
                Worksheet ws = (Worksheet)wb.Worksheets[1];

                // Write data to Excel
                int row = 1;
                foreach (var person in persons)
                {
                    ws.Cells[row, 1].Value = person.FirstName;
                    ws.Cells[row, 2].Value = person.Age;
                    ws.Cells[row, 3].Value = person.City;
                    row++;
                }

                // Save the workbook
                wb.SaveAs(filePath);
                wb.Close();

                Console.WriteLine("Data has been exported to Excel: " + filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while exporting data to Excel: " + ex.Message);
            }
            finally
            {
                // Quit Excel application
                excel.Quit();
            }
        }
    }
}
