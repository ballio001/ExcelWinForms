using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWinForm.Excel
{
    public class Person
    {
        public string FirstName { get; set; }
        public int Age { get; set; }
        public string City { get; set; }

        public Person(string firstName, int age, string city)
        {
            FirstName = firstName;
            Age = age;
            City = city;
        }
    }
}
