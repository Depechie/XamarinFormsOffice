using System;
using System.Collections.Generic;
using XamarinOfficeNetStandard.Models;

namespace XamarinOfficeNetStandard.Services
{
    public class DataService
    {
        public List<Person> GetEmployees()
        {
            List<Person> employees = new List<Person>();

            Company comp1 = new Company() { Name = "Company 1" };
            Company comp2 = new Company() { Name = "Company 2" };

            employees.Add(new Person()
            {
                FirstName = "Glenn",
                LastName = "Versweyveld",
                Email = "John@Do.com",
                Company = comp1
            });

            employees.Add(new Person()
            {
                FirstName = "Bart",
                LastName = "Lannoeye",
                Email = "John@Do.com",
                Company = comp1
            });

            employees.Add(new Person()
            {
                FirstName = "Jan",
                LastName = "Van de Poel",
                Email = "John@Do.com",
                Company = comp2
            });

            return employees;
        }
    }
}
