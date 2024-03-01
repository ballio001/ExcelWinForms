using ExcelWinForm.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ExcelWinForm.XML
{
    public interface IDataExporter
    {
        void ExportData(List<Person> persons);
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
            // Create XML document
            XDocument xmlDoc = new XDocument(
                new XElement("Persons",
                    from person in persons
                    select new XElement("Person",
                        new XElement("FirstName", person.FirstName),
                        new XElement("Age", person.Age),
                        new XElement("City", person.City)
                    )
                )
            );

            // Save XML document
            xmlDoc.Save(filePath);
            Console.WriteLine("Data has been exported to XML: " + filePath);
        }
    }
}