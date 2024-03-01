using System;
using System.Windows.Forms;

namespace ExcelWinForm
{
    public class Program
    {
        [STAThread]
        public static void Main()
        {        
            // Instantiate DatabaseManager with server name and database name
            DatabaseManager dbManager = new DatabaseManager("DESKTOP-2SAQ4OK\\MSSQLSERVER01", "WinForms");

            // Insert data from Excel into the database
            dbManager.InsertPersonsFromExcel();

            Application.Run(new Form1());
        }
    }
}
