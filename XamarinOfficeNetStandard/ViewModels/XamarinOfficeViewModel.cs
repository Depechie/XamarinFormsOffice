using System.Collections.Generic;
using System.Collections.ObjectModel;
using Xamarin.Forms;
using XamarinOfficeNetStandard.Helpers;
using XamarinOfficeNetStandard.Interfaces;
using XamarinOfficeNetStandard.Models;
using XamarinOfficeNetStandard.Services;

namespace XamarinOfficeNetStandard.ViewModels
{
    public class XamarinOfficeViewModel : BaseViewModel
    {
        private OpenXML _openXMl = new OpenXML();
        private DataService _dataService = new DataService();

        private ObservableCollection<Person> _employees;
        public ObservableCollection<Person> Employees
        {
            get => _employees;
            set
            {
                _employees = value;
                OnPropertyChanged();
            }
        }

        private Command _generateExcelCommand;
        public Command GenerateExcelCommand => _generateExcelCommand ?? (_generateExcelCommand = new Command(GenerateExcel));

        public XamarinOfficeViewModel()
        {
        }

        public override void LoadData()
        {
            base.LoadData();
            Employees = new ObservableCollection<Person>(_dataService.GetEmployees());
        }

        private void GenerateExcel()
        {
            string excelDocumentPath = _openXMl.GenerateExcel("TestExcel.xlsx");

            ExcelData data = new ExcelData();
            data.Headers.Add("First name");
            data.Headers.Add("Last name");
            data.Headers.Add("email");

            foreach(Person employee in Employees)
                data.Values.Add(new List<string>() { employee.FirstName, employee.LastName, employee.Email });

            _openXMl.InsertDataIntoSheet(excelDocumentPath, "Employees", data);

            DependencyService.Get<IShare>().Share(excelDocumentPath);
        }
    }
}