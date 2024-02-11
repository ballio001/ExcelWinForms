using System;
using System.Windows.Forms;
using ExcelWinForm.Excel;

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
            string excelData = ReadExcel.ReadExcelData();
            MessageBox.Show(excelData, "Excel Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void CmdWrite_Click(object sender, EventArgs e)
        {
            // Example data to write to Excel, TODO: add SQL data retrieval
            string[,] data = {
                { "Name", "Age", "City" },
                { "John", "30", "New York" },
                { "Alice", "25", "Los Angeles" },
                { "Bob", "35", "Chicago" }
            };

            WriteExcel.WriteExcelData(data);
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
