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

        public PosterHeaderDataTemplateViewModel(string posterName)
        {
            MyPosterName = posterName;
        }

    }
}
