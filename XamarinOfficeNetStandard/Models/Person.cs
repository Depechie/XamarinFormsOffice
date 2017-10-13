using System;
namespace XamarinOfficeNetStandard.Models
{
    public class Company
    {
        public string Name { get; set; }
    }

    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }

        public Company Company { get; set; }
    }
}
