using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace ExcelWinForm.Excel
{
    public class ReadExcel
    {
        public static List<Person> ReadExcelData()
        {
            string filePath = Files.EditedFilePath;
            Application excel = new Application();

            List<Person> persons = new List<Person>();

            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];

            Microsoft.Office.Interop.Excel.Range usedRange = ws.UsedRange;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                string firstName = ws.Cells[i, 1].Value != null ? ws.Cells[i, 1].Value.ToString() : "";

                int age;
                bool isAgeValid = int.TryParse(ws.Cells[i, 2].Value != null ? ws.Cells[i, 2].Value.ToString() : "", out age);

                string city = ws.Cells[i, 3].Value != null ? ws.Cells[i, 3].Value.ToString() : "";

                if (isAgeValid)
                {
                    Person person = new Person(firstName, age, city);
                    persons.Add(person);
                }
                else
                {
                    // Handle invalid age (e.g., log, display a message, skip the row, etc.)
                    Console.WriteLine($"Invalid age value at row {i}. Skipping the row.");
                }
            }

            // Close the workbook and Excel application
            wb.Close();
            excel.Quit();

            return persons;
        }
    }
}
