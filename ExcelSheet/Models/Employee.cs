using System;
using System.ComponentModel.DataAnnotations;

namespace ExcelSheet.Models
{
    public class Employee
    {
        public int ID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        [DataType(DataType.Date)]
        public DateTime HiringDate { get; set; }
        public int age { get; set; }
    }
}
