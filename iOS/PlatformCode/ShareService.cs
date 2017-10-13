using System;
using System.IO;
using System.Linq;
using XamarinOfficeNetStandard.Interfaces;
using CoreGraphics;
using Foundation;
using UIKit;

[assembly: Xamarin.Forms.Dependency(typeof(XamarinOffice.iOS.PlatformCode.ShareService))]
namespace XamarinOffice.iOS.PlatformCode
{
    public class ShareService : IShare
    {
        private UIDocumentInteractionController _controller;

        public void Share(string filePath)
        {
            UIApplication.SharedApplication.InvokeOnMainThread(() =>
            {
                _controller = UIDocumentInteractionController.FromUrl(NSUrl.FromFilename(filePath));
                _controller.Name = Path.GetFileName(filePath);
                var window = UIApplication.SharedApplication.KeyWindow;
                var subviews = window.Subviews;
                var view = subviews.Last();
                var frame = view.Frame;
                frame = new CGRect((float)Math.Min(10, frame.Width), (float)frame.Bottom, 0, 0);
                _controller.PresentOptionsMenu(frame, view, true);
            });            
        }
    }
}
