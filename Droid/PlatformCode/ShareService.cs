using System.IO;
using Android.App;
using Android.Content;
using XamarinOfficeNetStandard.Interfaces;
using Uri = Android.Net.Uri;

[assembly: Xamarin.Forms.Dependency(typeof(XamarinOffice.Droid.PlatformCode.ShareService))]
namespace XamarinOffice.Droid.PlatformCode
{
    public class ShareService : IShare
    {
        public void Share(string filePath)
        {
            Java.IO.File file = new Java.IO.File(filePath);

            Intent intent = new Intent(Intent.ActionView);
            string mimeType = string.Empty;

            if (Path.GetExtension(filePath).ToLower() == ".pdf")
                mimeType = "application/pdf";
            else if (Path.GetExtension(filePath).ToLower() == ".doc")
                mimeType = "application/msword";
            else if (Path.GetExtension(filePath).ToLower() == ".docx")
                mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            else if (Path.GetExtension(filePath).ToLower() == ".xls")
                mimeType = "application/vnd.ms-excel";
            else if (Path.GetExtension(filePath).ToLower() == ".xlsx")
                mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            else if (Path.GetExtension(filePath).ToLower() == ".ppt")
                mimeType = "application/vnd.ms-powerpoint";
            else if (Path.GetExtension(filePath).ToLower() == ".jpg")
                mimeType = "image/jpeg";

            var t = Uri.FromFile(file);
            intent.SetDataAndType(t, mimeType);

            intent.SetFlags(ActivityFlags.ClearWhenTaskReset | ActivityFlags.NewTask);
            this.StartActivity(intent);
        }
    }

    public static class ObjectExtensions
    {
        public static void StartActivity(this object o, Intent intent)
        {
            var context = o as Context;
            if (context != null)
                context.StartActivity(intent);
            else
            {
                intent.SetFlags(ActivityFlags.NewTask);
                Application.Context.StartActivity(intent);
            }
        }
    }
}
