using ExcelSheet.Models;
using IronXL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ExcelSheet.Services
{
    public class WriteOnExcel
    {
        public void write()
        {
            // create folder to save file
            string path1 = @"C:\IronXL\";
            string path2 = Path.Combine(path1, "temp1");

            // Create directory temp1 if it doesn't exist
            Directory.CreateDirectory(path2);


            char[] letters =
            {
                'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','Y','Z'
            };


            Type type = typeof(Employee);
            var propertyInfos = type.GetProperties();
            int attributeCount = 0;
            foreach (PropertyInfo property in type.GetProperties())
            {
                attributeCount += 1;
            }

            List<Employee> employees = new List<Employee>
            {
                new Employee()
                {

                ID = 1,
                FirstName = "bahy",
                LastName = "Talaat",
                HiringDate = DateTime.Now,
                age=24,
                },
                new Employee()
                {

                ID = 2,
                FirstName = "Mena",
                LastName = "Nageh",
                HiringDate = DateTime.Now,
                age=23,
                },
                 new Employee()
                {

                ID = 3,
                FirstName = "Bassam",
                LastName = "Nageh",
                HiringDate = DateTime.Now,
                age=24,
                },
                 new Employee()
                {

                ID = 4,
                FirstName = "Joseph",
                LastName = "Nageh",
                HiringDate = DateTime.Now,
                age=23,
                },
            };

           
             
            WorkBook wbook=null;
            try
            {
                wbook = WorkBook.Load(@"C:\IronXL\temp1\Test.xlsx");

            }
            catch(Exception ex)
            {
            }
            WorkSheet ws;
            if (wbook == null)
            {
                wbook = WorkBook.Create();
                ws = wbook.CreateWorkSheet("sheet1");


            }
            else
            {
                ws = wbook.GetWorkSheet("sheet1");
            }
            for (int j = 0; j < attributeCount; j++)
            {
                ws[letters[j].ToString() + 1].Value = propertyInfos[j].Name;

            }

            for (int i = 0; i < employees.Count; i++)
            {
                for (int j = 0; j < attributeCount; j++)
                {
                    ws[letters[j].ToString() + (i + 2)].Value = employees[i].GetType().GetProperty(propertyInfos[j].Name).GetValue(employees[i], null);
                }
            }
            wbook.SaveAs(@"C:\IronXL\temp1\Test.xlsx");

        }
    }
}
