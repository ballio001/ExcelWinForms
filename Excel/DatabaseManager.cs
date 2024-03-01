using ExcelWinForm.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ExcelWinForm
{
    public class DatabaseManager
    {
        private string connectionString;

        public DatabaseManager(string serverName, string databaseName)
        {
            // Create connection string
            connectionString = $"Data Source={serverName};Initial Catalog={databaseName};Integrated Security=True";
        }
        public void InsertPersonsFromExcel()
        {
            // Read data from Excel
            List<Person> persons = ReadExcel.ReadExcelData();

            // Insert each person into the database
            InsertPersons(persons);
        }

        public List<Person> GetPersonsFromDatabase()
        {
            List<Person> persons = new List<Person>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "SELECT FirstName, Age, City FROM Person";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        string firstName = reader["FirstName"].ToString();
                        int age = Convert.ToInt32(reader["Age"]);
                        string city = reader["City"].ToString();

                        persons.Add(new Person(firstName, age, city));
                    }

                    reader.Close();
                }
            }

            return persons;
        }

        private void InsertPersons(List<Person> persons)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                foreach (var person in persons)
                {
                    string query = "INSERT INTO Person (FirstName, Age, City) VALUES (@FirstName, @Age, @City)";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@FirstName", person.FirstName);
                        command.Parameters.AddWithValue("@Age", person.Age);
                        command.Parameters.AddWithValue("@City", person.City);

                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}

namespace ExcelWinForm.Excel
{
    public interface IDataExporter
    {
        void ExportData(List<Person> persons);
    }

    public class ExcelExporter : IDataExporter
    {
        private string filePath;

        public ExcelExporter(string filePath)
        {
            this.filePath = filePath;
        }

        public void ExportData(List<Person> persons)
        {
            // Code to export data to Excel
        }
    }

    public class XmlExporter : IDataExporter
    {
        private string filePath;

        public XmlExporter(string filePath)
        {
            this.filePath = filePath;
        }

        public void ExportData(List<Person> persons)
        {
            // Code to export data to XML
        }
    }
}