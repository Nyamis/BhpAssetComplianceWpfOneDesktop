using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using BhpAssetComplianceWpfOneDesktop.Constants;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels.DataTemplate
{
    public class RibbonButtonDataTemplateViewModel : BindableBase
    {
        private readonly Action _navigationAction;

        private string _myHeader;
        public string MyHeader
        {
            get { return _myHeader; }
            set { SetProperty(ref _myHeader, value); }
        }

        private ImageSource _myImage;
        public ImageSource MyImage
        {
            get { return _myImage; }
            set { SetProperty(ref _myImage, value); }
        }
        public ICommand PushCommand { get; set; }

        public RibbonButtonDataTemplateViewModel(string header, Action navigationAction, string relativeIconPath = "")
        {
            _navigationAction = navigationAction;
            PushCommand = new DelegateCommand(navigationAction);
            MyHeader = header;
         
            var directoryName = System.IO.Path.GetDirectoryName
                (System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            var iconPath = $"{directoryName}{relativeIconPath}";
            MyImage = new BitmapImage(new Uri(iconPath));
        }
    }
}
