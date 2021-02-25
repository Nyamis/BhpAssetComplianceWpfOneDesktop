using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Prism.Mvvm;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels.DataTemplate
{
    public class PosterHeaderDataTemplateViewModel : BindableBase
    {
        private string _myPosterName;
        public string MyPosterName
        {
            get { return _myPosterName; }
            set { SetProperty(ref _myPosterName, value); }
        }

        private ImageSource _myIcon;
        public ImageSource MyIcon
        {
            get { return _myIcon; }
            set { SetProperty(ref _myIcon, value); }
        }
        
        public PosterHeaderDataTemplateViewModel(string posterName,string posterIcon)
        {
            MyPosterName = posterName;
            var directoryName = System.IO.Path.GetDirectoryName
                (System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            var iconPath = $"{directoryName}{posterIcon}";
            MyIcon = new BitmapImage(new Uri(iconPath));
        }

    }
}
