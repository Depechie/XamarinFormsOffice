using Xamarin.Forms;
using XamarinOfficeNetStandard.ViewModels;

namespace XamarinOffice.Views
{
    public partial class XamarinOfficePage : ContentPage
    {
        public XamarinOfficePage()
        {
            InitializeComponent();
            BindingContext = new XamarinOfficeViewModel();
        }

        protected override void OnAppearing()
        {
            base.OnAppearing();
            ((BaseViewModel)BindingContext).LoadData();
        }
    }
}
