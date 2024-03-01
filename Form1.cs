using System;
using System.Windows.Forms;
using ExcelWinForm.Excel;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace ExcelWinForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void CmdRead_Click(object sender, EventArgs e)
        {
            try
            {
                // Fetch data from the SQL database
                DatabaseManager dbManager = new DatabaseManager("DESKTOP-2SAQ4OK\\MSSQLSERVER01", "WinForms");
                List<Person> persons = dbManager.GetPersonsFromDatabase();

                // Display the fetched data
                string message = "Data fetched from SQL database:\n\n";
                foreach (var person in persons)
                {
                    message += $"First Name: {person.FirstName}, Age: {person.Age}, City: {person.City}\n";
                }

                MessageBox.Show(message, "Data from SQL", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CmdWrite_Click(object sender, EventArgs e)
        {
            // Example data to write to Excel, TODO: add SQL data retrieval
            string[,] data = {
                        { "Id", "FirstName", "Age", "City" },
                        { "1", "John", "30", "New York" },
                        { "2", "Alice", "25", "Los Angeles" },
                        { "3", "Bob", "35", "Chicago" }
    };

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            saveFileDialog.Title = "Save Excel File";
            saveFileDialog.FileName = "ExcelOutput.csv";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                WriteExcel.WriteExcelData(data, filePath);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void CmdPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fileDialog.FileName;
            }
        }
    }
}
